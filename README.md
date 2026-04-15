# Keka → Microsoft 365 user sync

Nightly sync of employee attributes from Keka HR to Microsoft 365 (Entra ID) via Microsoft Graph.

**Match key:** Keka `email` == M365 `userPrincipalName`. Users in Keka but not in M365 are skipped (logged). Users in M365 but not in Keka are left untouched.

**Fields synced:** `jobTitle`, `department`, `officeLocation`, `city`, `country`, `mobilePhone`, `businessPhones`, `employeeId`, `employeeHireDate`, `manager`.

**Never touched:** `displayName`, `userPrincipalName`, `mail`, `accountEnabled`, licenses, group membership.

---

## One-time setup

### 1. Keka API credentials
1. In Keka: **Settings → API → Generate API Key** (requires Keka admin).
2. Note `client_id`, `client_secret`, `api_key`, and your tenant URL (e.g. `https://sarasanalytics.keka.com`).

### 2. Entra app registration
1. Entra admin center → **App registrations → New registration**. Name: `keka-m365-sync`. Single tenant.
2. **API permissions → Add → Microsoft Graph → Application permissions**:
   - `User.ReadWrite.All`
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
python sync.py --dry-run --verbose         # show what would change for everyone
python sync.py --apply --only you@sarasanalytics.com   # test on one user
python sync.py --apply                     # full sync
```

In CI:
- **Scheduled runs (daily 08:00 IST)** are forced to dry-run — check the workflow logs.
- **Manual runs:** Actions → "Keka → M365 sync" → **Run workflow** → tick `apply` to write.

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
