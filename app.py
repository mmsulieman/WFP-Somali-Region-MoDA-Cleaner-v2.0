import io
import re
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="WFP Jijiga AO – Process Monitoring Cleaner & KPI Summary (v5)",
    layout="wide"
)

# ---------- Helper functions ----------

SYSTEM_PREFIXES = ["_", "meta:", "instance", "start", "end"]
SYSTEM_NAMES = {
    "_id","_uuid","_submission_time","_xform_id","_version","_status",
    "starttime","endtime","deviceid","today"
}

def dedupe_columns(cols):
    """Ensure column names are unique by appending suffixes when needed."""
    seen = {}
    new_cols = []
    for c in cols:
        if c not in seen:
            seen[c] = 0
            new_cols.append(c)
        else:
            seen[c] += 1
            new_cols.append(f"{c}_{seen[c]}")
    return new_cols

def is_system_column(col: str) -> bool:
    low = col.lower().strip()
    if low in SYSTEM_NAMES:
        return True
    for p in SYSTEM_PREFIXES:
        if low.startswith(p):
            return True
    return False

def standardize_column_name(col: str) -> str:
    low = col.lower().strip()
    # location/admin
    if "region" in low and "sub office" not in low:
        return "region"
    if "sub office" in low:
        return "sub_office"
    if "zone" in low:
        return "zone"
    if "wereda" in low or "woreda" in low:
        return "woreda"
    if "kebele" in low and "specify" not in low:
        return "kebele"
    if "village" in low or ("camp" in low and "name" in low):
        return "village"
    if "fdp" in low and "name" in low:
        return "fdp_name"
    if "market" in low and "name" in low:
        return "market_name"
    if "health facility" in low or "tsfp centre" in low or "tsfp center" in low:
        return "tsfp_center"
    if "implementing partner" in low or "cooperating partner" in low or ("partner" in low and "name" in low):
        return "partner_name"
    # date / time
    if "date" in low and ("visit" in low or "monitor" in low or "interview" in low or low == "date"):
        return "visit_date"
    if "monitoring year" in low:
        return "monitoring_year"
    if "monitoring month" in low:
        return "monitoring_month"
    # enumerator
    if "information collected by" in low or "name of wfp staff" in low or "name of enumerator" in low or "interviewer" in low:
        return "enumerator"
    # gps
    if "latitude" in low:
        return "gps_lat"
    if "longitude" in low:
        return "gps_lon"
    if "altitude" in low:
        return "gps_alt"
    # default: cleaned snake_case
    cleaned = re.sub(r"[^0-9a-zA-Z]+","_", low).strip("_")
    return cleaned

def classify_meta_columns(columns):
    meta = []
    indicator = []
    for c in columns:
        low = c.lower()
        if any(p in low for p in ["region","zone","wereda","woreda","kebele","village","camp","fdp","market","tsfp","partner","date","month","year","enumerator","latitude","longitude","altitude"]):
            meta.append(c)
        else:
            indicator.append(c)
    return meta, indicator

def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Cleaned")
    buf.seek(0)
    return buf.getvalue()

# ---------- Sidebar / Help ----------

with st.sidebar:
    st.title("ℹ️ Help / Information")
    st.markdown("**WFP Jijiga Area Office – M&E Unit**")
    st.markdown(
        "- Upload raw MoDA / Excel / CSV process monitoring exports.\n"
        "- Page 1: Clean & standardize, create Tableau-ready long format.\n"
        "- Page 2: Explore basic KPI summaries for quick QA."
    )
    with st.expander("Cleaning logic", expanded=False):
        st.markdown(
            "- Drops Kobo/system columns (e.g. `_id`, `_uuid`, `_submission_time`, `starttime`, `endtime`).\n"
            "- Deduplicates column names with suffixes `_1`, `_2`, etc. when needed (both before and after renaming).\n"
            "- Standardizes location fields to: `region`, `sub_office`, `zone`, `woreda`, `kebele`, `village`, `fdp_name`, `tsfp_center`, `market_name`.\n"
            "- Standardizes partner/date fields: `partner_name`, `visit_date`, `monitoring_year`, `monitoring_month`.\n"
            "- Long format: melts all non-metadata fields into `indicator_name` / `indicator_value`."
        )

st.markdown("<h1 style='color:#0072BC;'>WFP Jijiga – Process Monitoring Cleaner & KPI Summary (v5)</h1>", unsafe_allow_html=True)
st.caption("Standalone app for MoDA/Excel process monitoring data – cleaning, deduplicating, reshaping, and KPI snapshots.")

uploaded = st.file_uploader(
    "📁 Upload MoDA / Excel / CSV Process Monitoring File",
    type=["xlsx","xls","csv"]
)

if uploaded is not None:
    st.success(f"Uploaded: {uploaded.name}")
    filetype = "csv" if uploaded.name.lower().endswith(".csv") else "excel"

    # ----- Load and deduplicate columns safely -----
    if filetype == "excel":
        xls = pd.ExcelFile(uploaded)
        sheet_name = st.selectbox("Select sheet to process", xls.sheet_names)
        df_raw = pd.read_excel(xls, sheet_name=sheet_name)
    else:
        sheet_name = None
        df_raw = pd.read_csv(uploaded)

    # Handle duplicate column names before anything else
    duplicate_flag = df_raw.columns.duplicated().any()
    if duplicate_flag:
        df_raw.columns = dedupe_columns(df_raw.columns)
        st.warning("Duplicate column names were found in the raw file and have been automatically renamed with suffixes (e.g. '_1', '_2').")

    # Use tabs as two "pages"
    tab_clean, tab_kpi = st.tabs(["🧼 Cleaning & Tableau Prep", "📊 Basic KPI Summaries"])

    # ---------- TAB 1: Cleaning & Tableau Prep ----------
    with tab_clean:
        drop_system = st.checkbox("Drop Kobo/system fields (recommended)", value=True, key="drop_system")
        do_standardize = st.checkbox("Standardize key metadata fields (region, woreda, partner, date, etc.)", value=True, key="std_cols")
        make_long = st.checkbox("Generate Tableau-ready long-format dataset", value=True, key="make_long")

        log_lines = []

        df = df_raw.copy()

        # Safety: dedupe again before preview
        if df.columns.duplicated().any():
            df.columns = dedupe_columns(df.columns)

        st.write("### 🔍 Raw Preview (after deduplicating column names if needed)")
        st.dataframe(df.head())

        # Drop system columns
        if drop_system:
            before = df.shape[1]
            keep_cols = [c for c in df.columns if not is_system_column(c)]
            dropped = [c for c in df.columns if c not in keep_cols]
            df = df[keep_cols]
            log_lines.append(f"Dropped {before - len(keep_cols)} system columns: {', '.join(dropped) if dropped else 'None'}")

        # Standardize names
        if do_standardize:
            new_cols = {c: standardize_column_name(c) for c in df.columns}
            df.rename(columns=new_cols, inplace=True)
            # Safety: dedupe AGAIN after renaming, in case different raw names map to same standardized name
            if df.columns.duplicated().any():
                df.columns = dedupe_columns(df.columns)
                log_lines.append("Detected duplicate column names after standardization; renamed with numeric suffixes to keep them unique.")
            else:
                log_lines.append("Standardized column names to snake_case and aligned key metadata fields (region, woreda, partner_name, etc.).")

        # Final safety: ensure no duplicates before display
        if df.columns.duplicated().any():
            df.columns = dedupe_columns(df.columns)

        st.write("### ✅ Cleaned (Wide) Preview")
        st.dataframe(df.head())

        # Decide meta vs indicator columns
        meta_cols, indicator_cols = classify_meta_columns(df.columns.tolist())
        for c in ["region","sub_office","zone","woreda","kebele","village","fdp_name","tsfp_center","market_name","partner_name","visit_date","enumerator"]:
            if c in df.columns and c not in meta_cols:
                meta_cols.append(c)
                if c in indicator_cols:
                    indicator_cols.remove(c)
        meta_cols = sorted(list(dict.fromkeys(meta_cols)))  # unique & stable

        st.write("#### 🧩 Detected metadata fields")
        st.code(", ".join(meta_cols) if meta_cols else "None detected – the app will treat all fields as indicators for long-format.")

        df_long = None
        if make_long:
            if not meta_cols:
                st.warning("No metadata fields detected – long-format will use all columns as indicators.")
                meta_use = []
            else:
                meta_use = meta_cols

            indicator_use = [c for c in df.columns if c not in meta_use]
            log_lines.append(f"Identified {len(meta_use)} metadata columns and {len(indicator_use)} indicator columns for long-format reshaping.")

            if indicator_use:
                df_long = df.melt(
                    id_vars=meta_use,
                    value_vars=indicator_use,
                    var_name="indicator_name",
                    value_name="indicator_value"
                )
                # Safety: long-format has fixed column names; no duplicates expected, but check anyway
                if df_long.columns.duplicated().any():
                    df_long.columns = dedupe_columns(df_long.columns)

                st.write("### 📊 Tableau-ready Long Format Preview")
                st.dataframe(df_long.head())
            else:
                st.warning("No indicator columns identified for long-format dataset.")

        # Downloads
        st.markdown("### 💾 Downloads")

        # Wide cleaned Excel
        wide_bytes = to_excel_bytes(df)
        wide_name = "PM_Cleaned_Wide.xlsx" if sheet_name is None else f"PM_Cleaned_Wide_{sheet_name}.xlsx"
        st.download_button(
            "⬇️ Download Cleaned Wide File (Excel)",
            data=wide_bytes,
            file_name=wide_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Long-format CSV
        if make_long and df_long is not None:
            long_csv = df_long.to_csv(index=False).encode("utf-8")
            long_name = "PM_LongFormat_Tableau.csv" if sheet_name is None else f"PM_LongFormat_Tableau_{sheet_name}.csv"
            st.download_button(
                "⬇️ Download Tableau-ready Long Format (CSV)",
                data=long_csv,
                file_name=long_name,
                mime="text/csv"
            )

        # Logs
        st.write("### 📝 Processing Log")
        if duplicate_flag:
            log_lines.insert(0, "Detected duplicate column names in the input; renamed with numeric suffixes to make them unique.")
        if log_lines:
            st.write("\n".join(f"- {ln}" for ln in log_lines))
        else:
            st.write("No transformations applied.")

        # Store cleaned data & long format in session_state for KPI tab
        st.session_state["clean_wide"] = df
        st.session_state["meta_cols"] = meta_cols
        if df_long is not None:
            st.session_state["long_df"] = df_long

    # ---------- TAB 2: Basic KPI Summaries ----------
    with tab_kpi:
        st.write("### 📊 Basic KPI Summaries & QA Checks")

        if "long_df" not in st.session_state:
            st.info("First go to **Cleaning & Tableau Prep** tab, enable 'Generate Tableau-ready long-format dataset', then return here.")
        else:
            df_long = st.session_state["long_df"]
            meta_cols = st.session_state.get("meta_cols", [])

            # Safety: make sure no duplicate column names
            if df_long.columns.duplicated().any():
                df_long.columns = dedupe_columns(df_long.columns)

            st.write("#### Dataset overview")
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total records (rows)", f"{len(df_long):,}")
            with col2:
                st.metric("Unique indicators", df_long["indicator_name"].nunique())
            with col3:
                meta_count = len(meta_cols) if meta_cols else 0
                st.metric("Metadata fields", meta_count)

            # Filters for location / partner if available
            filter_cols = []
            for c in ["sub_office","woreda","fdp_name","tsfp_center","market_name","partner_name"]:
                if c in df_long.columns:
                    filter_cols.append(c)

            if filter_cols:
                with st.expander("Filter by geography / partner", expanded=False):
                    for c in filter_cols:
                        vals = sorted([v for v in df_long[c].dropna().unique()])
                        if len(vals) > 0:
                            selected = st.multiselect(f"{c} filter", vals, default=vals)
                            if selected:
                                df_long = df_long[df_long[c].isin(selected)]

            # KPI selection
            st.write("#### Choose an indicator (KPI) to explore")
            if "indicator_name" not in df_long.columns:
                st.warning("No 'indicator_name' column found in the long-format dataset.")
            else:
                indicator_options = sorted(df_long["indicator_name"].dropna().unique())
                if not indicator_options:
                    st.warning("No indicator_name values found in the long-format dataset.")
                else:
                    selected_indicator = st.selectbox("Indicator (from indicator_name)", indicator_options)
                    df_kpi = df_long[df_long["indicator_name"] == selected_indicator].copy()

                    # Try to convert to numeric for stats
                    df_kpi["numeric_value"] = pd.to_numeric(df_kpi["indicator_value"], errors="coerce")

                    kpi_col1, kpi_col2, kpi_col3 = st.columns(3)
                    with kpi_col1:
                        st.metric("N (rows for this KPI)", f"{len(df_kpi):,}")
                    with kpi_col2:
                        st.metric("Non-null values", df_kpi["indicator_value"].notna().sum())
                    with kpi_col3:
                        if df_kpi["numeric_value"].notna().any():
                            mean_val = df_kpi["numeric_value"].mean()
                            st.metric("Mean (numeric)", f"{mean_val:0.3f}")
                        else:
                            st.metric("Mean (numeric)", "N/A")

                    # Visualization: bar distribution for categorical / histogram-like for numeric
                    st.write("#### Distribution for selected KPI")
                    if df_kpi["numeric_value"].notna().any():
                        st.bar_chart(df_kpi["numeric_value"])
                    else:
                        freq = (
                            df_kpi["indicator_value"]
                            .fillna("Missing")
                            .value_counts()
                            .reset_index()
                            .rename(columns={"index": "category","indicator_value":"count"})
                        )
                        st.dataframe(freq)
                        st.bar_chart(freq.set_index("category"))

else:
    st.info("Upload a MoDA/Excel/CSV file to begin cleaning, reshaping, and exploring KPIs.")

st.markdown("""---
<div style="font-size:0.9rem; color:#555;">
<strong>WFP Jijiga Area Office – M&E Unit (RAM)</strong><br>
Process Monitoring Data Cleaning, Tableau Preparation & KPI Snapshot Utility – Streamlit App v5
</div>
""", unsafe_allow_html=True)
