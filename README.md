# Karad Division — COD Digital Transaction % Report

Streamlit app to generate the COD (Cash on Delivery) Digital Transaction %
report for Karad Division, India Post.

**Live app:** https://karad-cod-digital-report.streamlit.app *(update after deployment)*

---

## How to use (daily workflow)

1. Open the app
2. Set dates:
   - **Consolidated period — From date** — defaults to 1st of current month
   - **Consolidated period — To date** — defaults to yesterday
   - **Single date** — defaults to yesterday
3. Drop the 2 CSV files in their respective slots (file names don't matter):
   - ① COD Collection — Consolidated period CSV
   - ② COD Collection — Single date CSV
4. Click **Generate COD Digital Transaction Report**
5. Review the summary metrics → click **Download**

---

## Files in this repo

| File                      | Purpose                                                          |
| ------------------------- | ----------------------------------------------------------------- |
| `app.py`                  | Streamlit UI                                                      |
| `report_builder.py`       | All Excel generation logic                                        |
| `Office_Master_File.xlsx` | Office master (288 offices) — update here when structure changes  |
| `requirements.txt`        | Python dependencies                                                |
| `.streamlit/config.toml`  | Theme                                                              |

---

## What the report contains

A single `.xlsx` with 3 sheets:

1. **Raw Data** — every office, two date blocks (consolidated + single date),
   Total COD / Digital COD / Digital Txn % each.
2. **Sub Division wise Summary** — Table A (one row per Sub Office, B.O.s rolled
   up into their parent S.O./H.O.) and Table B below it (Sub Division totals +
   Karad Division grand total), volume-weighted.
3. **Summary** — plain-language narrative: COD received / delivered digitally /
   Digital Txn % per Sub Division plus a Division total, for both date ranges,
   with auto-generated remarks (Healthy / Moderate / Low digital adoption).

All percentages render as `xx.xx%`. Offices with zero COD activity show blank
cells rather than `0%`. Offices with COD volume but zero digital uptake show
`0.00%` in bold red. No panes are frozen on any sheet.

---

## Built-in data-quality check

Every run, the app checks whether each Sub Office's own row agrees with its
Branch Offices on which Sub Division it belongs to. If a Sub Office's parent
row is tagged differently from its unanimous B.O.s, the app auto-corrects it
(using the B.O. majority as truth) and reports the correction on-screen. If
a Sub Office name turns out to be a genuine collision between two unrelated
sub-offices in different divisions, the app keeps them separate (grouped by
Sub Division + Sub Office Name together) and flags this instead of merging
them.

---

## Updating the master file

When office structure changes (new offices, discontinued offices):

1. Prepare the updated `Office_Master_File.xlsx`
2. Upload it to this repo (replace the existing file)
3. The app picks it up automatically on next run (cached for 1 hour)

OR use the **Upload new master file** button in the app sidebar — but note
that this only applies for the current session. For it to persist, you still
need to push to the repo.

---

## Deployment (Streamlit Cloud)

1. Fork / push this repo to `dokaradmmu/karad-cod-digital-report`
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. New app → select this repo → `main` branch → `app.py`
4. Deploy

No secrets or environment variables needed.
