# Split File Combiner

A Streamlit app that combines "split" files that share the same header/template
— e.g. a large export chopped into `1.xlsx`, `2.xlsx`, `3.xlsx`, or any set of
files where the top few rows repeat and only the data underneath differs.

Originally built for Shopee's Mass Update exports, but it's generic: works
with any `.xlsx`, `.xls`, or `.csv` files (can even mix types), whether
zipped together or uploaded individually.

## What it does

- Accepts a zip, several individual files, or a mix of both
- Auto-detects how many leading rows are the shared header/template by
  comparing rows across your files — no hardcoded row count
- Lets you override the detected header row count if it guesses wrong
- Shows a preview of the first file's rows so you can sanity-check before combining
- Auto-patches Shopee's broken `activePane` XML attribute (harmless no-op for
  files that don't have that issue)
- Skips any file that's an empty template (header only, no data rows) and
  reports it separately
- Gives you one combined `.xlsx` to download, preserving the base file's
  formatting/validation where possible

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open the local URL Streamlit prints (usually `http://localhost:8501`).

## Deploy for free (Streamlit Community Cloud)

1. Push this repo to your GitHub account.
2. Go to [share.streamlit.io](https://share.streamlit.io) and sign in with GitHub.
3. Click **New app**, pick this repo/branch, set the main file to `app.py`.
4. Deploy — you'll get a shareable URL you can open from your phone or any browser.

## Project structure

```
app.py                # Streamlit UI
combiner_engine.py     # Core combine logic (reusable, no Streamlit dependency)
requirements.txt
```

## Notes on header detection

If you upload 2+ files, header rows are auto-detected as the longest run of
leading rows that are identical (cell-by-cell, whitespace-trimmed) across
every file. If you only upload one file, there's nothing to compare against,
so it defaults to 1 header row — adjust the number input in the app if your
file actually has more.
