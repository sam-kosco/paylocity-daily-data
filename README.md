# paylocity-daily-data

Automated nightly sync of Paylocity scheduling and labor hours data into SharePoint for Foxtrot Aviation Services. Runs on GitHub Actions — no local machine required.

---

## What It Does

### Schedule Sheet
Pulls all scheduled shifts for every active employee and writes them to the **Schedule** tab of `Schedule and Hours.xlsx` on SharePoint.

- **Range:** 2 weeks back through 2 weeks forward, recalculated from today on every run
- **Updates:** Additions, modifications, and cancellations from Paylocity are all reflected — the entire update window is replaced each run
- **History:** Rows older than the 14-day update window are preserved from prior runs and never re-queried

### Labor Hours Sheet
Pulls actual worked punch data and writes aggregated daily totals to the **Labor Hours** tab.

- **Range:** 60 days back, accumulating across runs
- **Update window:** Only the last 14 days are re-queried each run to conserve API calls — older rows are left as-is
- **Aggregation:** All punches for a given employee on a given date are collapsed into a single row (total hours + total estimated earnings)
- **Timezone:** All dates and times are in **EST**

---

## Output File

**SharePoint:** `Data Hub → Shared Documents → Definitive Lists → Schedule and Hours.xlsx`

### Schedule tab columns

| Column | Description |
|---|---|
| Shift ID | Unique Paylocity shift identifier |
| Employee ID | Paylocity employee ID |
| Employee Name | First Last |
| Start DateTime | Shift start in EST (YYYY-MM-DD HH:MM) |
| End DateTime | Shift end in EST (YYYY-MM-DD HH:MM) |
| Labor Distribution | Cost Center 0 code |
| Work Scope | Cost Center 1 code |
| Scheduled Time | Shift duration in hours |

### Labor Hours tab columns

| Column | Description |
|---|---|
| Employee ID | Paylocity employee ID |
| Employee Name | First Last |
| Date | Work date (YYYY-MM-DD, EST) |
| Hours Worked | Total hours from all ClockOut punches that day |
| Labor Distribution | Cost Center 0 code |
| Work Scope | Cost Center 1 code |
| Earnings | Gross estimated earnings (Paylocity-calculated) |

---

## Schedule

Runs nightly at **2:00 AM EST** via GitHub Actions cron.

Can also be triggered manually: **Actions → Schedule & Hours Sync → Run workflow**

---

## Setup

### 1. Add GitHub Secrets

Go to **Settings → Secrets and variables → Actions** and add the following. These secrets are **not shared** between repos — they must be added here even if they exist in `ramp-refresh` or `stl-report-automation`.

| Secret | Value |
|---|---|
| `PAYLOCITY_CLIENT_ID` | Paylocity API client ID |
| `PAYLOCITY_CLIENT_SECRET` | Paylocity API client secret |
| `TENANT_ID` | `ede0c57f-549f-4a90-9f8c-7ea130346f95` |
| `CLIENT_ID` | `58191600-ab56-4141-bff6-806805fcbff4` (Foxtrot Report Automation) |
| `CLIENT_SECRET` | Entra app client secret — same credential used in `ramp-refresh` and `stl-report-automation` |

### 2. Confirm Paylocity Scheduling API Access

The scheduling endpoint (`/v2/.../scheduling/shifts`) may need to be explicitly enabled on your Paylocity account. Confirm with your Paylocity representative that your API credentials have access before the first run.

### 3. Place Files

```
paylocity-daily-data/
├── schedule_hours_sync.py
└── .github/
    └── workflows/
        └── schedule_hours_sync.yml
```

---

## Troubleshooting

**Check the run log:** Actions → Schedule & Hours Sync → click the failed run → expand **Run sync**

| Error | Likely Cause | Fix |
|---|---|---|
| `401 Unauthorized` (Paylocity) | API credentials expired or wrong | Regenerate in Paylocity → update `PAYLOCITY_CLIENT_ID` / `PAYLOCITY_CLIENT_SECRET` |
| `401 Unauthorized` (Graph API) | `CLIENT_SECRET` expired | Renew in Entra → App registrations → Foxtrot Report Automation → Certificates & secrets → update `CLIENT_SECRET` in **all three repos** |
| `404` on employee shifts | Scheduling API not enabled for account | Contact Paylocity rep to enable scheduling API access |
| Workbook not updating | Graph API permissions missing | Verify `Files.ReadWrite.All` is granted in Entra for Foxtrot Report Automation |
| Token expired mid-run | Run took >1 hour | Script refreshes Paylocity token every 50 employees — if still failing, reduce batch size |

---

## Credential Maintenance

| Credential | Expires | Action |
|---|---|---|
| `CLIENT_SECRET` (Entra) | Every 24 months | Renew in Entra → update in `ramp-refresh`, `stl-report-automation`, and **this repo** |
| `PAYLOCITY_CLIENT_SECRET` | Every 12 months | Self-renew via `POST /api/v2/credentials/secrets` using renewal code emailed by Paylocity → update here |

> ⚠️ Set a calendar reminder 23 months after creating the Entra `CLIENT_SECRET`. When it expires, it breaks **all three automation repos** simultaneously.

---

## Related Repos

| Repo | What it does |
|---|---|
| [ramp-refresh](https://github.com/sam-kosco/ramp-refresh) | Ramp card transaction sync to SharePoint |
| [stl-report-automation](https://github.com/sam-kosco/stl-report-automation) | STL AA Cabin nightly PDF report |

All three repos share the same **Foxtrot Report Automation** Entra app and the same `CLIENT_SECRET`.

---

*Foxtrot Aviation Services · Data & Analytics*
