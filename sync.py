"""
Keka HR → Microsoft 365 (Entra ID) user attribute sync.

Pulls employees from Keka and patches matching M365 users via Microsoft Graph.
Match key: Keka `email` == M365 `userPrincipalName`.

Run modes:
  python sync.py --inspect           # dump one Keka employee JSON, no writes
  python sync.py --dry-run           # show diffs, no writes (default in CI without APPLY=true)
  python sync.py --apply             # actually patch users
  python sync.py --apply --only user@example.com   # patch a single user

Requires env vars (see .env.example):
  KEKA_BASE_URL, KEKA_CLIENT_ID, KEKA_CLIENT_SECRET, KEKA_API_KEY
  AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import sys
import time
from typing import Any, Callable

import requests
from msal import ConfidentialClientApplication

# Optional: load .env for local runs. CI uses real env vars, so this is a no-op there.
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# ---------- Configuration ----------

# Keka groupType taxonomy (confirmed from sarasanalytics tenant response):
#   1 = Business Unit, 2 = Department, 3 = Location, 4 = Cost Center,
#   5 = Pay Group, 9 = Legal Entity
def _group(emp: dict, group_type: int) -> str | None:
    for g in (emp.get("groups") or []):
        if g.get("groupType") == group_type:
            return g.get("title")
    return None

# Map Keka employee object → Graph user property.
# Each entry is (graph_property, extractor_function).
# Returning None or "" means "skip this field for this user" (won't overwrite).
FIELD_MAP: list[tuple[str, Callable[[dict], Any]]] = [
    ("jobTitle",         lambda e: _g(e, "jobTitle", "title")),
    ("department",       lambda e: _group(e, 2)),
    ("officeLocation",   lambda e: _group(e, 3)),
    ("city",             lambda e: e.get("city")),
    ("country",          lambda e: e.get("countryCode")),  # 2-letter ISO; Graph accepts
    ("mobilePhone",      lambda e: e.get("mobilePhone")),
    ("businessPhones",   lambda e: [e["workPhone"]] if e.get("workPhone") else None),
    ("employeeId",       lambda e: e.get("employeeNumber")),
    ("employeeHireDate", lambda e: _date(e.get("joiningDate"))),
]

# Manager handled separately (needs $ref to manager's Graph object id).
MANAGER_KEY = lambda e: _g(e, "reportsTo", "email")

# Only sync active employees (Keka accountStatus: 1 = active, others = inactive/exited).
def is_active(emp: dict) -> bool:
    return emp.get("accountStatus") == 1 and not emp.get("exitDate")

# Don't touch users whose UPN matches these patterns (svc accounts, shared mailboxes).
SKIP_UPN_PATTERNS = ("svc-", "shared-", "noreply", "admin@")

GRAPH = "https://graph.microsoft.com/v1.0"

# ---------- Helpers ----------

def _g(d: dict | None, *keys) -> Any:
    """Safe nested getter."""
    cur = d
    for k in keys:
        if not isinstance(cur, dict):
            return None
        cur = cur.get(k)
    return cur

def _date(v: Any) -> str | None:
    """Coerce Keka date (often ISO 8601 with time) to Graph's date-only format."""
    if not v:
        return None
    return str(v)[:10]  # YYYY-MM-DD

def _setup_logging(verbose: bool) -> None:
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(asctime)s %(levelname)-7s %(message)s",
        datefmt="%H:%M:%S",
    )

# ---------- Keka ----------

KEKA_AUTH_URL = "https://login.keka.com/connect/token"
# Keka sits behind Azure App Gateway which blocks default python-requests UA.
KEKA_HEADERS = {"User-Agent": "keka-m365-sync/1.0", "Accept": "application/json"}

def keka_token() -> str:
    r = requests.post(
        KEKA_AUTH_URL,
        data={
            "grant_type": "kekaapi",
            "scope": "kekaapi",
            "client_id": os.environ["KEKA_CLIENT_ID"],
            "client_secret": os.environ["KEKA_CLIENT_SECRET"],
            "api_key": os.environ["KEKA_API_KEY"],
        },
        headers=KEKA_HEADERS,
        timeout=30,
    )
    if not r.ok:
        raise RuntimeError(f"Keka auth {r.status_code}: {r.text[:300]}")
    return r.json()["access_token"]

def keka_employees(token: str) -> list[dict]:
    """Fetch all active employees, paginated."""
    base = os.environ["KEKA_BASE_URL"].rstrip("/")
    headers = {**KEKA_HEADERS, "Authorization": f"Bearer {token}"}
    all_emps: list[dict] = []
    page = 1
    while True:
        r = requests.get(
            f"{base}/api/v1/hris/employees",
            headers=headers,
            params={"pageNumber": page, "pageSize": 200},
            timeout=60,
        )
        if not r.ok:
            raise RuntimeError(f"Keka employees {r.status_code}: {r.text[:300]}")
        body = r.json()
        batch = body.get("data", body if isinstance(body, list) else [])
        if not batch:
            break
        all_emps.extend(batch)
        # Stop when fewer than a full page returned.
        if len(batch) < 200:
            break
        page += 1
    logging.info("Keka: fetched %d employees", len(all_emps))
    return all_emps

# ---------- Graph ----------

def graph_token() -> str:
    app = ConfidentialClientApplication(
        client_id=os.environ["AZURE_CLIENT_ID"],
        client_credential=os.environ["AZURE_CLIENT_SECRET"],
        authority=f"https://login.microsoftonline.com/{os.environ['AZURE_TENANT_ID']}",
    )
    res = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in res:
        raise RuntimeError(f"Graph token failed: {res.get('error_description')}")
    return res["access_token"]

def _graph_request(method: str, url: str, token: str, **kwargs) -> requests.Response:
    """Wrapper with throttling-aware retry."""
    headers = kwargs.pop("headers", {})
    headers["Authorization"] = f"Bearer {token}"
    for attempt in range(5):
        r = requests.request(method, url, headers=headers, timeout=30, **kwargs)
        if r.status_code == 429 or r.status_code >= 500:
            wait = int(r.headers.get("Retry-After", 2 ** attempt))
            logging.warning("Graph %s - sleeping %ds (attempt %d)", r.status_code, wait, attempt + 1)
            time.sleep(wait)
            continue
        return r
    return r  # type: ignore

def graph_get_user(token: str, upn: str) -> dict | None:
    props = "id,userPrincipalName,jobTitle,department,officeLocation,city,country,mobilePhone,businessPhones,employeeId,employeeHireDate,onPremisesSyncEnabled"
    r = _graph_request("GET", f"{GRAPH}/users/{upn}?$select={props}", token)
    if r.status_code == 404:
        return None
    r.raise_for_status()
    return r.json()

def graph_get_manager(token: str, upn: str) -> str | None:
    """Return current manager's UPN, or None."""
    r = _graph_request("GET", f"{GRAPH}/users/{upn}/manager?$select=userPrincipalName", token)
    if r.status_code == 404:
        return None
    r.raise_for_status()
    return r.json().get("userPrincipalName")

def graph_patch_user(token: str, upn: str, body: dict) -> None:
    r = _graph_request("PATCH", f"{GRAPH}/users/{upn}", token, json=body)
    if r.status_code >= 400:
        raise RuntimeError(f"PATCH {upn} failed {r.status_code}: {r.text}")

def graph_set_manager(token: str, upn: str, manager_id: str) -> None:
    r = _graph_request(
        "PUT",
        f"{GRAPH}/users/{upn}/manager/$ref",
        token,
        json={"@odata.id": f"{GRAPH}/users/{manager_id}"},
    )
    if r.status_code >= 400:
        raise RuntimeError(f"set manager {upn} failed {r.status_code}: {r.text}")

# ---------- Diff & apply ----------

def build_desired(emp: dict) -> dict:
    out: dict = {}
    for prop, extractor in FIELD_MAP:
        try:
            val = extractor(emp)
        except Exception as e:
            logging.debug("extract %s failed: %s", prop, e)
            val = None
        if val in (None, ""):
            continue
        out[prop] = val
    return out

def diff(desired: dict, current: dict) -> dict:
    """Return only fields that need updating."""
    changes = {}
    for k, v in desired.items():
        cur = current.get(k)
        # Graph returns employeeHireDate as full ISO datetime ("2026-01-05T00:00:00Z")
        # even though we PATCH it as YYYY-MM-DD. Normalize to date-only for comparison.
        if k == "employeeHireDate" and isinstance(cur, str):
            cur = cur[:10]
        if isinstance(v, list) and isinstance(cur, list):
            if sorted(map(str, v)) != sorted(map(str, cur)):
                changes[k] = v
        elif cur != v:
            changes[k] = v
    return changes

def should_skip(upn: str) -> bool:
    u = upn.lower()
    return any(p in u for p in SKIP_UPN_PATTERNS)

# ---------- Main ----------

def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--inspect", action="store_true", help="print one Keka employee record and exit")
    ap.add_argument("--dry-run", action="store_true", help="show diffs, do not write")
    ap.add_argument("--apply", action="store_true", help="actually write changes to M365")
    ap.add_argument("--only", help="restrict to one UPN (for testing)")
    ap.add_argument("--verbose", "-v", action="store_true")
    args = ap.parse_args()

    _setup_logging(args.verbose)

    # CI safety: APPLY=true env var also enables writes (paired with workflow_dispatch input).
    apply = args.apply or os.environ.get("APPLY", "").lower() == "true"
    if not apply and not args.dry_run and not args.inspect:
        logging.info("No mode given - defaulting to --dry-run (use --apply to write)")
        args.dry_run = True

    ktok = keka_token()
    employees = keka_employees(ktok)

    if args.inspect:
        if not employees:
            print("(no employees returned)")
            return 1
        print(json.dumps(employees[0], indent=2, default=str))
        return 0

    gtok = graph_token()

    # Build UPN → Graph object id map for manager resolution (1 lookup per Keka manager).
    manager_id_cache: dict[str, str | None] = {}
    def resolve_manager_id(upn: str) -> str | None:
        if upn in manager_id_cache:
            return manager_id_cache[upn]
        u = graph_get_user(gtok, upn)
        manager_id_cache[upn] = u["id"] if u else None
        return manager_id_cache[upn]

    n_match = n_changed = n_skipped = n_failed = 0
    n_mgr_changed = 0

    for emp in employees:
        upn = (emp.get("email") or "").strip()
        if not upn:
            continue
        if args.only and upn.lower() != args.only.lower():
            continue
        if not is_active(emp):
            logging.debug("inactive: %s", upn)
            continue
        if should_skip(upn):
            logging.debug("skip pattern: %s", upn)
            continue

        current = graph_get_user(gtok, upn)
        if current is None:
            logging.info("not in M365: %s", upn)
            n_skipped += 1
            continue
        n_match += 1

        if current.get("onPremisesSyncEnabled"):
            logging.warning("AD-synced user, Graph PATCH will be rejected - skipping: %s", upn)
            n_skipped += 1
            continue

        desired = build_desired(emp)
        changes = diff(desired, current)

        if changes:
            logging.info("CHANGE %s: %s", upn, json.dumps(changes, default=str))
            if apply:
                try:
                    graph_patch_user(gtok, upn, changes)
                    n_changed += 1
                except RuntimeError as e:
                    if "403" in str(e):
                        # User holds a privileged Entra role — service principal
                        # with User.ReadWrite.All cannot write admins. Expected;
                        # HR updates these manually. Don't fail the workflow.
                        logging.warning("SKIP (admin/no-permission): %s", upn)
                        n_skipped += 1
                    else:
                        logging.error("FAIL %s: %s", upn, e)
                        n_failed += 1
            else:
                n_changed += 1  # would-change

        # Manager (separate call; only if Keka has one)
        keka_mgr_upn = MANAGER_KEY(emp)
        if keka_mgr_upn:
            current_mgr = graph_get_manager(gtok, upn)
            if (current_mgr or "").lower() != keka_mgr_upn.lower():
                logging.info("MANAGER %s: %s → %s", upn, current_mgr, keka_mgr_upn)
                if apply:
                    mgr_id = resolve_manager_id(keka_mgr_upn)
                    if mgr_id:
                        try:
                            graph_set_manager(gtok, upn, mgr_id)
                            n_mgr_changed += 1
                        except RuntimeError as e:
                            if "403" in str(e):
                                logging.warning("SKIP manager (admin/no-permission): %s", upn)
                            else:
                                logging.error("FAIL manager %s: %s", upn, e)
                                n_failed += 1
                    else:
                        logging.warning("manager not in M365: %s", keka_mgr_upn)
                else:
                    n_mgr_changed += 1

    mode = "APPLIED" if apply else "DRY-RUN"
    logging.info(
        "%s - matched=%d, attr-changes=%d, manager-changes=%d, skipped=%d, failed=%d",
        mode, n_match, n_changed, n_mgr_changed, n_skipped, n_failed,
    )
    return 1 if n_failed else 0


if __name__ == "__main__":
    sys.exit(main())
