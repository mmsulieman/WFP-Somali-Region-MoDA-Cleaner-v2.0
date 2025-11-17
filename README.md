# WFP Somali Region – Process Monitoring Cleaner & KPI Summary (Streamlit v5)

This version adds **extra-safe handling of duplicate column names**, including after standardization,
which was causing PyArrow / Streamlit dataframe errors.

## Key features

- Upload raw **MoDA / Excel / CSV** process monitoring exports.
- Automatically:
  - Deduplicates column names at three stages:
    - Immediately after file load.
    - After column standardization.
    - Before any dataframe previews.
  - Removes Kobo/system fields (`_id`, `_uuid`, `_submission_time`, `starttime`, `endtime`, etc.).
  - Standardizes common metadata fields:
    - `region`, `sub_office`, `zone`, `woreda`, `kebele`, `village`
    - `fdp_name`, `tsfp_center`, `market_name`
    - `partner_name`, `visit_date`, `monitoring_year`, `monitoring_month`, `enumerator`
  - Splits metadata vs indicator fields and reshapes to:
    - **Wide cleaned file** (Excel)
    - **Long-format Tableau-ready file** (`indicator_name`, `indicator_value`)

- Provides **two pages (tabs)**:
  - **🧼 Cleaning & Tableau Prep** – for data preparation & export.
  - **📊 Basic KPI Summaries** – for QA and exploratory checks of individual KPIs.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open the URL shown in your terminal (usually http://localhost:8501).

## Recommended workflow

1. Export raw MoDA/ODK/Excel/CSV from the field monitoring system.
2. On **Cleaning & Tableau Prep**, upload the file, choose sheet (if Excel), and keep defaults:
   - ✅ Drop Kobo/system fields
   - ✅ Standardize key metadata fields
   - ✅ Generate long-format dataset
3. Download:
   - `PM_Cleaned_Wide*.xlsx` → archive, QA, Excel pivots.
   - `PM_LongFormat_Tableau*.csv` → connect to Tableau / Power BI (and join to your indicator metadata table).
4. Move to **Basic KPI Summaries** to sanity-check key indicators before loading into dashboards.
