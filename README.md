# CTG South Dashboard — Excel → HTML auto-update

`index.html` is the dashboard. Its data lives inside the file as `const NAME = …;` blocks.
When you upload a new Excel to `data/`, GitHub Actions rebuilds those blocks and commits the new `index.html`.

## Daily use (3 steps)
1. On GitHub open the repo → `data/` folder → delete the old `.xlsx` and **Add file → Upload files** the new one
   (keep **exactly one** `.xlsx` in `data/`; any file name is fine).
   Open/save it in Excel first so formula results are stored in the file.
2. Commit directly to `main`.
3. Open the **Actions** tab → "Update dashboard from Excel" turns green in ~1 minute → `index.html` is updated
   (GitHub Pages / your host picks it up automatically).

## One-time setup
- Push this whole folder to a GitHub repo (keep `.github/workflows/update-dashboard.yml`).
- Repo → Settings → Actions → General → Workflow permissions → **Read and write permissions**.
- (Optional) Settings → Pages → deploy from branch `main`, folder `/ (root)`.

## Run locally
```
pip install openpyxl
python scripts/build_dashboard.py --check   # dry run: which blocks would change
python scripts/build_dashboard.py           # rewrite index.html
```
See `HANDOFF.md` for how it works and how to continue from another account.
