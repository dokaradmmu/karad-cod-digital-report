import streamlit as st
import requests
import io
import re
from datetime import date, timedelta
from report_builder import build_report

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Karad Division — COD Digital Transaction Report",
    page_icon="📮",
    layout="wide",
)

MASTER_RAW_URL = (
    "https://raw.githubusercontent.com/dokaradmmu/karad-cod-digital-report/main/Office_Master_File.xlsx"
)

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown(
    """
    <div style="background:#1F3864;padding:18px 24px;border-radius:8px;margin-bottom:24px">
        <h2 style="color:white;margin:0;font-family:Arial">📮 Karad Division — COD Digital Transaction % Report</h2>
        <p style="color:#DCE6F1;margin:4px 0 0 0;font-size:14px">
            Office of the Superintendent of Post Offices, Karad Division
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

# ── Sidebar — master file management ─────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Master File")
    st.caption("Office_Master_File.xlsx is embedded in the GitHub repo and loaded automatically each run.")

    uploaded_master = st.file_uploader(
        "Upload new master file (replaces repo copy)",
        type=["xlsx"],
        key="master_upload",
        help="Only needed when office structure changes. Push the new file to the repo root to make it permanent."
    )

    if uploaded_master:
        st.success("New master file loaded for this session. Push it to the repo to make it permanent.")

    st.divider()
    st.markdown("**App info**")
    st.caption("Karad Division, Pune Region, Maharashtra Circle")
    st.caption("Digital Txn % = COD Digital Count ÷ Total COD Count")
    st.caption("Offices with no COD activity show blank cells, not 0%.")

# ── Load master file ──────────────────────────────────────────────────────────
@st.cache_data(ttl=3600, show_spinner=False)
def fetch_master_from_repo():
    resp = requests.get(MASTER_RAW_URL, timeout=15)
    resp.raise_for_status()
    return resp.content

def get_master_bytes():
    if uploaded_master:
        return uploaded_master.read()
    try:
        return fetch_master_from_repo()
    except Exception as e:
        st.error(f"❌ Could not load master file from GitHub: {e}")
        st.stop()

# ── Date inputs ───────────────────────────────────────────────────────────────
today = date.today()
yesterday = today - timedelta(days=1)
first_of_month = today.replace(day=1)

st.markdown("### 📅 Report Dates")
col1, col2, col3 = st.columns(3)
with col1:
    cons_from = st.date_input("Consolidated period — From date", value=first_of_month, format="DD/MM/YYYY")
with col2:
    cons_to = st.date_input("Consolidated period — To date", value=yesterday, format="DD/MM/YYYY")
with col3:
    single_date = st.date_input("Single date", value=yesterday, format="DD/MM/YYYY")

if cons_from > cons_to:
    st.error("❌ Consolidated period 'From date' cannot be after 'To date'.")
    st.stop()

consolidated_label = f"{cons_from.strftime('%d.%m.%Y')} to {cons_to.strftime('%d.%m.%Y')}"
single_date_label = single_date.strftime('%d.%m.%Y')
file_tag = single_date.strftime('%d_%m_%Y')

st.info(
    f"📋 Consolidated period: **{consolidated_label}**  |  Single date: **{single_date_label}**  |  "
    f"Filename: **Karad_COD_Digital_Transaction_Report_{file_tag}.xlsx**"
)

# ── File upload slots ─────────────────────────────────────────────────────────
st.markdown("---")
st.markdown("### 📁 Upload Files")
st.caption("File names don't matter — drop each file in its designated slot.")

col_a, col_b = st.columns(2)

with col_a:
    st.markdown(
        f"**① COD Collection — Consolidated period**  \n"
        f"<span style='color:#555;font-size:13px'>Period: {consolidated_label}</span>",
        unsafe_allow_html=True
    )
    cod_cons_file = st.file_uploader("", type=["csv"], key="cod_cons")

with col_b:
    st.markdown(
        f"**② COD Collection — Single date**  \n"
        f"<span style='color:#555;font-size:13px'>Date: {single_date_label}</span>",
        unsafe_allow_html=True
    )
    cod_single_file = st.file_uploader("", type=["csv"], key="cod_single")

# ── Generate ──────────────────────────────────────────────────────────────────
st.markdown("---")

all_uploaded = all([cod_cons_file, cod_single_file])

if not all_uploaded:
    missing = []
    if not cod_cons_file:
        missing.append("① Consolidated period COD")
    if not cod_single_file:
        missing.append("② Single date COD")
    st.warning(f"⚠️ Waiting for: {', '.join(missing)}")

generate_btn = st.button(
    "🚀 Generate COD Digital Transaction Report",
    disabled=not all_uploaded,
    type="primary",
    use_container_width=True,
)

if generate_btn and all_uploaded:
    with st.spinner("Building report…"):
        try:
            master_bytes = get_master_bytes()

            xlsx_bytes, totals, corrections = build_report(
                master_file=io.BytesIO(master_bytes),
                consolidated_csv=io.BytesIO(cod_cons_file.read()),
                single_date_csv=io.BytesIO(cod_single_file.read()),
                consolidated_label=consolidated_label,
                single_date_label=single_date_label,
                file_tag=file_tag,
            )

            # ── Data-quality corrections ─────────────────────────────────
            if corrections:
                st.warning(
                    "⚠️ **Data-quality corrections applied automatically:**\n\n"
                    + "\n\n".join(f"- {c}" for c in corrections)
                )
            else:
                st.success("✅ No Sub Division / Sub Office mismatches found in the master file.")

            # ── Summary metrics ──────────────────────────────────────────
            st.markdown("### ✅ Report Generated — Division Summary")

            col1, col2, col3 = st.columns(3)
            col1.metric("Consolidated — Total COD Articles", f"{totals['cons_total']:,}")
            col2.metric("Consolidated — Digital Articles", f"{totals['cons_digital']:,}")
            col3.metric(
                "Consolidated — Digital Txn %",
                f"{totals['cons_pct']:.2f}%" if totals['cons_pct'] is not None else "—"
            )

            col4, col5, col6 = st.columns(3)
            col4.metric("Single Date — Total COD Articles", f"{totals['single_total']:,}")
            col5.metric("Single Date — Digital Articles", f"{totals['single_digital']:,}")
            col6.metric(
                "Single Date — Digital Txn %",
                f"{totals['single_pct']:.2f}%" if totals['single_pct'] is not None else "—"
            )

            # ── Download ───────────────────────────────────────────────────
            filename = f"Karad_COD_Digital_Transaction_Report_{file_tag}.xlsx"
            st.download_button(
                label=f"⬇️ Download {filename}",
                data=xlsx_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )

        except Exception as e:
            st.error(f"❌ Error generating report: {e}")
            st.exception(e)
