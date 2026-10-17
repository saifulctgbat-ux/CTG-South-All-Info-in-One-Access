# HANDOFF — read this first if you are continuing this project (new account / new session)

## Goal
User uploads a new Excel (`data/*.xlsx`) to GitHub → GitHub Action rebuilds the data embedded in `index.html` → dashboard updates. No manual HTML editing.

## Status (as of 2026-10-17): DONE and verified
- `scripts/build_dashboard.py` — rebuilds all 13 embedded data blocks from the workbook. Idempotent (2nd run = "unchanged").
- `.github/workflows/update-dashboard.yml` — runs the script on push of `data/*.xlsx`, commits `index.html` back.
- `index.html` — already rebuilt from the current Excel. `data/CTG_South.xlsx` — current Excel.
- Verification done: every block was rebuilt from the Excel and compared with the original HTML; all matched except places where the OLD HTML was stale vs the Excel (Sep'26 numbers in `DATA`, ~20 YLP sub-channels, 2 KPI cells) — the Excel is treated as truth. Inline JS of the new HTML passes a syntax check.
- NOT yet done: pushing to the user's real GitHub repo (repo name was not given). User must push the folder / enable Actions write permission (see README).

## How the script works
`index.html` is both template and output. For each `const NAME = …;` it rebuilds the value and replaces only that statement (JSON blocks via `json.JSONDecoder.raw_decode`; `DIST_DATA`, `BA_SHEET_DATA`, `PW_DATA` are regenerated JS literals). CSS/JS untouched. Sheets are found by name prefix (`rows_of`), months/columns are discovered from header text, so adding rows or a month column usually just works. Safety: refuses to write if a block comes out nearly empty.

## Sheet → block map
| Excel sheet | HTML const | notes |
|---|---|---|
| TSA Wise By  Channel Outlet + TSA Wise Target Vs Achievement | `DATA` | joined on TSA Location; 1 dp; status = ach>=100; last item = CTG South total row (`location:0`) |
| TSA Wise Target Vs Achievement | `TSA_TVA_DATA` | 2 dp, ach in %, status text copied from sheet; months: Feb'26, March'26, April'26, May'26, June'26, July'26, Aug'26, Sep'26 |
| Distribution Overview | `DIST_DATA` | BAT/AKTC/JTI/PMI |
| TSA Wise Partner Winner Status | `PW_DATA` | only "Winner Number" columns |
| CTG South Base 2026 | `BA_SHEET_DATA` | sections Area / DH / Territory / Distribution Point, 2 dp |
| Price List 2026 | `PL_ROWS` | |
| YLP Point Jan-June'26 | `YLP_RAW` | T/L/R = first-seen lookup lists, D = rows by index |
| KPI(Partner,TSA,TMC) | `KPI_RECORDS_SOURCE` | sheet type "CM Sup KPI" is called "MS KPI" in HTML (`TYPE_RENAME`); empty own cells -> "—" |
| ByPartner Target Vs Achievement | `BP_RECORDS_SOURCE`, `BP_MEMO_BY_MONTH`, `BP_MONTH_BRAND_CONFIG`, `BP_MEMO_SPLIT_AUG`, `BP_SUPER_OVER` | month keys here are Jan'26…Sep'26 (3 letters). Memo columns in records = first brand (Jan–Jul), B&H+LS sum (Aug), blank (Sep) — copied from how the original HTML behaved |

## Known limitations / next ideas
- Hard-coded in the HTML's JS (not data): `PW_MONTHS`, `YLP_MONTHS`/`YLP_MONTH_COLS`, `AUG_TILL_DATE`/`SEP_TILL_DATE`, `bpSuperOver` only for Sep'26, `BP` Oct'26 not present. A NEW month (e.g. Oct'26 in the ByPartner/TVA sheets) is picked up in the data, but the UI may need JS changes to show it.
- If the ByPartner sheet gets a new month, check `bp_groups()` header parsing and the Aug/Sep special cases in `build_bp()`.
- Excel must be saved by Excel (cached formula values). If a formula cell has no cached value the script falls back to computing pct/status for ByPartner only.
- `index.html` is ~18 MB; each auto-commit adds that to repo history (fine, GitHub limit is 100 MB per file).

## Quick commands
```
pip install openpyxl
python scripts/build_dashboard.py --check      # dry run
python scripts/build_dashboard.py              # build in place
python scripts/build_dashboard.py --xlsx X --html Y --out Z
```
## If you are a new Claude session
Upload this folder (or the repo), ask: "Read HANDOFF.md and continue". Remaining work is only: user's GitHub push/permissions, and any UI changes for new months.
