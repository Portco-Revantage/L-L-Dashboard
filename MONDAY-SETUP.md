# Connect the dashboard to Monday.com

Every 30 minutes GitHub reads your Monday board, updates `data.xlsx` and republishes the dashboard.

| What comes from Monday | What stays in Excel |
|---|---|
| Projects sheet (fully refreshed every sync) | Portco details (industry, SPOC, email) |
| New portco names (added to the Portcos sheet) | Communications you log by hand |
| Item updates (added to Communications as "Monday update") | Files sheet |

## Files to add to your GitHub repo

| File | Where |
|---|---|
| `monday_sync.py` | top level |
| `monday-config.json` | top level |
| `app.js`, `index.html`, `style.css` | top level (replace the current ones) |
| `monday-sync.yml` | must be at `.github/workflows/monday-sync.yml` |

## Monday column names it recognises (not case-sensitive)

| Dashboard field | Monday column titles |
|---|---|
| Project name | the item name |
| Portco | Portco, Client, Company (otherwise the **group name** is used) |
| Owner | ID Owner, Owner, Person, Assignee, Designer |
| Status | Status, Stage |
| Due date | Due Date, Due, Deadline, Timeline (end date), Date |
| Start date | Start Date, Start (or Timeline start) |
| Script received | Script Received, Script Date, Request Date |
| Progress | Progress, % Complete (otherwise worked out from the status) |
| Feedback rounds | Feedback Rounds, Feedback, Review Rounds |
| Content type / Priority | Content Type, Type, Format / Priority |

Monday status labels are translated, e.g. **Done → Delivered**, **Working on it → In Development**,
**Stuck → On Hold**. Add your own labels to `status_map` in `monday-config.json`.

## Safety

- The Monday token is stored as a GitHub **secret**: it is never visible in the code or on the dashboard.
- If Monday returns no projects (e.g. wrong board ID), `data.xlsx` is **not** changed.
- If the token is missing, the sync is skipped and the dashboard keeps its current data.
- In a **public** repo, the synced `data.xlsx` is visible to anyone. Use a private repo for real data.
