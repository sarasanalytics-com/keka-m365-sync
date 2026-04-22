# Keka → Microsoft 365 user sync

Nightly sync of employee attributes from Keka HR to Microsoft 365 (Entra ID) via Microsoft Graph.

**Match key:** Keka `email` == M365 `userPrincipalName`. Users in Keka but not in M365 are skipped (logged). Users in M365 but not in Keka are left untouched.

**Fields synced:** `jobTitle`, `department`, `officeLocation`, `city`, `country`, `mobilePhone`, `businessPhones`, `employeeId`, `employeeHireDate`, `manager`.

**Never touched:** `displayName`, `userPrincipalName`, `mail`, `accountEnabled`, licenses.

Optionally syncs **Distribution List memberships** — see [DL sync](#distribution-list-dl-sync) below.

---

## One-time setup

### 1. Keka API credentials
1. In Keka: **Settings → API → Generate API Key** (requires Keka admin).
2. Note `client_id`, `client_secret`, `api_key`, and your tenant URL (e.g. `https://sarasanalytics.keka.com`).

### 2. Entra app registration
1. Entra admin center → **App registrations → New registration**. Name: `keka-m365-sync`. Single tenant.
2. **API permissions → Add → Microsoft Graph → Application permissions**:
   - `User.ReadWrite.All`
   - `GroupMember.ReadWrite.All` *(only required if using DL sync)*
3. Click **Grant admin consent**.
4. **Certificates & secrets → New client secret**. Copy the *value* (only shown once).
5. From **Overview** copy: Tenant ID, Application (client) ID.

### 3. GitHub secrets
In the repo: **Settings → Secrets and variables → Actions → New repository secret**. Add all 7:

| Secret | From |
|---|---|
| `KEKA_BASE_URL` | e.g. `https://sarasanalytics.keka.com` |
| `KEKA_CLIENT_ID` | Keka |
| `KEKA_CLIENT_SECRET` | Keka |
| `KEKA_API_KEY` | Keka |
| `AZURE_TENANT_ID` | Entra app overview |
| `AZURE_CLIENT_ID` | Entra app overview |
| `AZURE_CLIENT_SECRET` | Entra app secret value |

---

## Verifying the field map (do this first)

Keka's response shape varies by tenant config. Before trusting the sync, dump one employee record and confirm the field paths in `FIELD_MAP` (top of `sync.py`) match.

Locally:
```bash
cp .env.example .env   # fill in values
python -m venv .venv && source .venv/bin/activate   # or .venv\Scripts\activate on Windows
pip install -r requirements.txt
python sync.py --inspect
```

Look at the printed JSON. If e.g. `department.name` is actually `departmentInfo.displayName`, edit `FIELD_MAP` accordingly.

---

## Running

```bash
python sync.py --dry-run --verbose                         # show what would change for everyone
python sync.py --apply --only you@sarasanalytics.com       # test on one user
python sync.py --apply                                     # full attribute sync
python sync.py --apply --dl-sync                           # attributes + DL memberships
python sync.py --dry-run --dl-sync --verbose               # preview DL changes only (no writes)
python sync.py --apply --dl-sync --dl-rules custom.json    # use a different rules file
```

In CI:
- **Scheduled runs (weekly Sunday 08:00 IST)** are forced to dry-run — check the workflow logs.
- **Manual runs:** Actions → "Keka → M365 sync" → **Run workflow** → tick `apply` and/or `dl_sync` to write.

---

## Distribution List (DL) sync

The `--dl-sync` flag reconciles **who belongs to each M365 Distribution Group** based on Keka employee data.

### How it works

1. Reads rules from `dl_rules.json` (you create this from `dl_rules.example.json`).
2. For each rule, fetches the current member list from M365.
3. Calculates who should be added (matches filter, not yet a member) and who should be removed (is a Keka employee who no longer matches the filter).
4. **External contacts, guests, and non-Keka accounts already in a DL are never removed.** Only Keka-known employees are managed.
5. Additions are sent in batches of 20 (Graph API limit).

### Setup

**Step 1 — Entra permission.** Add `GroupMember.ReadWrite.All` application permission to your app registration and grant admin consent (see [Entra app registration](#2-entra-app-registration) above).

**Step 2 — Create `dl_rules.json`.**

```bash
cp dl_rules.example.json dl_rules.json
# edit dl_rules.json to match your DLs and Keka group names
```

Each rule maps a DL email to a Keka filter:

```json
[
  {
    "dl_email": "all-staff@sarasanalytics.com",
    "filter": "all"
  },
  {
    "dl_email": "engineering@sarasanalytics.com",
    "filter": { "field": "department", "value": "Engineering" }
  },
  {
    "dl_email": "bangalore@sarasanalytics.com",
    "filter": { "field": "location", "value": "Bangalore" }
  }
]
```

**Supported `filter` forms:**

| `filter` value | Matches employees where… |
|---|---|
| `"all"` | every active employee |
| `{"field": "department", "value": "X"}` | Keka Department == X |
| `{"field": "location", "value": "X"}` | Keka Location == X |
| `{"field": "business_unit", "value": "X"}` | Keka Business Unit == X |
| `{"field": "cost_center", "value": "X"}` | Keka Cost Center == X |

Values are matched case-insensitively. Use `python sync.py --inspect` to see the exact Keka group names for your tenant.

**Step 3 — Dry-run first.**

```bash
python sync.py --dry-run --dl-sync --verbose
```

This shows every planned addition and removal without writing anything.

**Step 4 — Apply.**

```bash
python sync.py --apply --dl-sync
```

### Troubleshooting DL sync

| Symptom | Cause |
|---|---|
| `DL not found in M365: x@y.com` | The `dl_email` doesn't match the group's primary SMTP address in M365 |
| `add members to group ... 403` | Missing `GroupMember.ReadWrite.All` permission or admin consent not granted |
| `fetch members failed: 403` | Same as above |
| Members not removed after department change | Check the `value` in the rule exactly matches the Keka group name (run `--inspect`) |

---

## Operational notes

- **Outlook GAL refresh:** changes appear in users' Outlook within ~24h (offline address book cycle). The user object itself updates immediately.
- **Manager loop / circular references:** Graph rejects cycles. Script logs and continues.
- **Service accounts:** UPNs containing `svc-`, `shared-`, `noreply`, `admin@` are skipped (see `SKIP_UPN_PATTERNS`).
- **AD-synced users:** if any user has `onPremisesSyncEnabled=true`, the script logs a warning and skips — Graph rejects PATCH on synced objects.
- **Audit:** all PATCH operations are visible in **Entra → Monitoring → Audit logs** (filter by service principal = the app name).

## Troubleshooting

| Symptom | Likely cause |
|---|---|
| `Graph token failed: AADSTS7000215` | wrong client secret |
| `Graph token failed: AADSTS65001` | admin consent not granted |
| All users return 404 | UPNs in Keka don't match M365 (e.g. `@sarasanalytics.com` vs `@saras-analytics.com`) |
| `PATCH ... 403 Insufficient privileges` | missing `User.ReadWrite.All` application permission |
| `extract X failed` in logs | Keka field name differs from `FIELD_MAP` — re-run `--inspect` |
