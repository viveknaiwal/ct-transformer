import io
import re
from datetime import datetime
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="CT Data Transformer",
    page_icon="🔄",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&display=swap');

html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }

/* Page background */
.stApp { background: #f7f5f2; }
.block-container { padding: 2.5rem 2rem 4rem 2rem; max-width: 1100px; }

/* Hide default streamlit elements */
#MainMenu, footer, header { visibility: hidden; }
.stDeployButton { display: none; }

/* Hero */
.hero { 
    background: #1a1714; 
    border-radius: 8px; 
    padding: 40px 44px 36px;
    margin-bottom: 28px;
    position: relative;
    overflow: hidden;
}
.hero::before {
    content: '';
    position: absolute;
    top: -60px; right: -60px;
    width: 280px; height: 280px;
    background: radial-gradient(circle, rgba(212,82,26,0.25) 0%, transparent 70%);
    pointer-events: none;
}
.hero-label {
    font-size: 11px; font-weight: 600; letter-spacing: 0.12em;
    text-transform: uppercase; color: #d4521a;
    display: flex; align-items: center; gap: 8px; margin-bottom: 14px;
}
.hero-label::before { content: ''; width: 20px; height: 2px; background: #d4521a; }
.hero-title { font-size: 36px; font-weight: 300; color: #f5f0eb; letter-spacing: -0.02em; margin-bottom: 8px; }
.hero-title strong { font-weight: 700; }
.hero-sub { font-size: 14px; color: #8a7f75; line-height: 1.6; max-width: 520px; }

/* Metric cards */
.metric-row { display: flex; gap: 12px; margin-bottom: 24px; }
.metric-card {
    flex: 1; background: #fff; border: 1px solid #e5e1db;
    border-radius: 6px; padding: 18px 20px;
    border-left: 3px solid #d4521a;
}
.metric-val { font-size: 26px; font-weight: 700; font-family: 'DM Mono', monospace; color: #d4521a; line-height: 1; }
.metric-label { font-size: 10px; font-weight: 600; letter-spacing: 0.08em; text-transform: uppercase; color: #8a7f75; margin-top: 4px; }

/* Section headers */
.section-head {
    font-size: 10px; font-weight: 600; letter-spacing: 0.1em;
    text-transform: uppercase; color: #8a7f75;
    display: flex; align-items: center; gap: 8px;
    margin-bottom: 14px; padding-bottom: 10px;
    border-bottom: 1px solid #e5e1db;
}
.section-head span { background: #d4521a; color: #fff; font-size: 9px; padding: 2px 7px; border-radius: 2px; }

/* Upload area */
[data-testid="stFileUploadDropzone"] {
    background: #fff !important;
    border: 2px dashed #d4c8bc !important;
    border-radius: 6px !important;
    transition: all 0.2s !important;
}
[data-testid="stFileUploadDropzone"]:hover {
    border-color: #d4521a !important;
    background: rgba(212,82,26,0.03) !important;
}

/* Selectboxes */
[data-testid="stSelectbox"] > div > div {
    background: #fff !important;
    border: 1px solid #e5e1db !important;
    border-radius: 4px !important;
    font-size: 13px !important;
}

/* Buttons */
.stButton > button, .stDownloadButton > button {
    border-radius: 4px !important;
    font-family: 'DM Sans', sans-serif !important;
    font-weight: 600 !important;
    font-size: 13px !important;
    letter-spacing: 0.02em !important;
    transition: all 0.2s !important;
}
.stButton > button[kind="primary"], .stDownloadButton > button {
    background: #d4521a !important;
    border: none !important;
    color: #fff !important;
}
.stButton > button[kind="primary"]:hover, .stDownloadButton > button:hover {
    background: #c4420a !important;
    transform: translateY(-1px) !important;
    box-shadow: 0 4px 16px rgba(212,82,26,0.3) !important;
}

/* Success / info / error */
.stSuccess { background: rgba(42,122,75,0.07) !important; border-color: rgba(42,122,75,0.25) !important; border-radius: 4px !important; }
.stInfo    { background: rgba(212,82,26,0.06) !important; border-color: rgba(212,82,26,0.2) !important; border-radius: 4px !important; }
.stError   { border-radius: 4px !important; }
.stWarning { border-radius: 4px !important; }

/* Dataframe */
[data-testid="stDataFrame"] { border: 1px solid #e5e1db !important; border-radius: 6px !important; overflow: hidden; }

/* Checkboxes */
[data-testid="stCheckbox"] label { font-size: 13px !important; color: #3a3530 !important; }

/* Divider */
hr { border-color: #e5e1db !important; }

/* Expander */
[data-testid="stExpander"] {
    background: #fff !important;
    border: 1px solid #e5e1db !important;
    border-radius: 6px !important;
}
</style>
""", unsafe_allow_html=True)

# ── Column candidates ─────────────────────────────────────────────────────────
COLUMN_CANDIDATES = {
    "emp_id":       ["emp id", "empid", "employee id", "ecode", "emp_id", "employee_code"],
    "emp_name":     ["employee name", "name", "emp name", "employee"],
    "when_change":  ["when was the change", "when_was_the_change", "effective date",
                     "when changed", "change date", "whenchange"],
    "total_ctc":    ["total ct", "total ctc", "ctc", "total_ct", "total"],
    "created_date": ["created date", "created_date", "created on", "created",
                     "created_at", "createdat"]
}

# ── Helpers ───────────────────────────────────────────────────────────────────
def find_column(columns, candidates):
    cols_map = {c.lower().strip(): c for c in columns}
    for cand in candidates:
        if cand.lower().strip() in cols_map:
            return cols_map[cand.lower().strip()]
    for col_lower, orig in cols_map.items():
        for cand in candidates:
            if cand.lower().strip() in col_lower:
                return orig
    return None

def try_parse_dates(series, dayfirst=True):
    parsed = pd.to_datetime(series, dayfirst=dayfirst, errors="coerce")
    if parsed.isna().mean() > 0.2:
        alt = pd.to_datetime(series, dayfirst=not dayfirst, errors="coerce")
        if alt.isna().mean() < parsed.isna().mean():
            parsed = alt
    return parsed

def clean_ctc(x):
    if pd.isna(x): return pd.NA
    if isinstance(x, (int, float)): return x
    s = re.sub(r"[^\d\.\-]", "", str(x).strip())
    try: return float(s)
    except: return pd.NA

def to_excel_bytes(df):
    # BUG FIX: use BytesIO correctly — must call getvalue() before context exits
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="CT Pivoted")
        ws = writer.sheets["CT Pivoted"]
        for col in ws.columns:
            max_len = max((len(str(cell.value)) if cell.value else 0) for cell in col)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 3, 30)
    return buf.getvalue()

def safe_index(columns, value):
    if value is None: return 0
    cols = list(columns)
    try: return cols.index(value) + 1
    except: return 0

# ── Hero ──────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero">
  <div class="hero-label">Payroll Internal Tool</div>
  <div class="hero-title">CT Data <strong>Transformer</strong></div>
  <div class="hero-sub">Upload your employee CT file. The tool keeps the latest record per change date and pivots the data — one row per employee, each change date as a column.</div>
</div>
""", unsafe_allow_html=True)

# ── Step 1: Upload ────────────────────────────────────────────────────────────
st.markdown('<div class="section-head"><span>01</span> Upload File</div>', unsafe_allow_html=True)

uploaded = st.file_uploader(
    "Drop your CSV or Excel file here",
    type=["csv", "xlsx", "xls"],
    accept_multiple_files=False,
    label_visibility="collapsed"
)

if uploaded is None:
    st.info("📂 Upload a file above to get started.")
    st.stop()

# ── Read file ─────────────────────────────────────────────────────────────────
try:
    raw_bytes = uploaded.read()
    if uploaded.name.lower().endswith(".csv"):
        df = pd.read_csv(io.BytesIO(raw_bytes), dtype=str)
    else:
        df = pd.read_excel(io.BytesIO(raw_bytes), dtype=str)
except Exception as e:
    st.error(f"❌ Failed to read file: {e}")
    st.stop()

st.success(f"✅ **{uploaded.name}** loaded — {len(df):,} rows · {len(df.columns)} columns")

st.divider()

# ── Step 2: Options + Column mapping ─────────────────────────────────────────
st.markdown('<div class="section-head"><span>02</span> Configure</div>', unsafe_allow_html=True)

col_opt1, col_opt2, col_opt3 = st.columns(3)
with col_opt1:
    dayfirst = st.checkbox("Dates are DD-MM-YYYY (day first)", value=True)
with col_opt2:
    auto_clean = st.checkbox("Clean Total CT (remove commas/currency)", value=True)
with col_opt3:
    drop_empty = st.checkbox("Drop all-empty columns in output", value=True)

fc1, fc2 = st.columns([1, 3])
with fc1:
    filter_date = st.date_input(
        "Filter: remove Change Dates before",
        value=pd.Timestamp("2024-01-01"),
        min_value=pd.Timestamp("2000-01-01"),
        max_value=pd.Timestamp("2030-12-31"),
        help="Any row where 'When was the change?' is before this date will be excluded"
    )

st.markdown("**Column Mapping** — auto-detected from your headers:")

detected = {k: find_column(df.columns.tolist(), v) for k, v in COLUMN_CANDIDATES.items()}
options_with_none = [None] + list(df.columns)

mc1, mc2, mc3, mc4, mc5 = st.columns(5)
with mc1: emp_col    = st.selectbox("EMP ID",           options_with_none, index=safe_index(df.columns, detected["emp_id"]))
with mc2: name_col   = st.selectbox("Employee Name",    options_with_none, index=safe_index(df.columns, detected["emp_name"]))
with mc3: when_col   = st.selectbox("When Was Change?", options_with_none, index=safe_index(df.columns, detected["when_change"]))
with mc4: ct_col     = st.selectbox("Total CT",         options_with_none, index=safe_index(df.columns, detected["total_ctc"]))
with mc5: created_col= st.selectbox("Created Date",     options_with_none, index=safe_index(df.columns, detected["created_date"]))

missing = [k for k,v in {"EMP ID":emp_col,"Employee Name":name_col,
                          "When Was Change?":when_col,"Total CT":ct_col,
                          "Created Date":created_col}.items() if not v]
if missing:
    st.error(f"❌ Please map these columns: **{', '.join(missing)}**")
    st.stop()

st.divider()

# ── Step 3: Process ───────────────────────────────────────────────────────────
st.markdown('<div class="section-head"><span>03</span> Process</div>', unsafe_allow_html=True)

if st.button("⚡  Process & Generate Output", type="primary", use_container_width=True):
    with st.spinner("Processing..."):

        # Select & rename
        work = df[[emp_col, name_col, when_col, ct_col, created_col]].copy()
        work.columns = ["EMP_ID","EMP_NAME","WHEN_CHG","TOTAL_CT","CREATED"]

        # Parse dates
        work["WHEN_DT"]    = try_parse_dates(work["WHEN_CHG"], dayfirst=dayfirst)
        work["CREATED_DT"] = try_parse_dates(work["CREATED"],  dayfirst=dayfirst)

        # ── FILTER: drop rows where Change Date is before cutoff ──────────
        cutoff = pd.Timestamp(filter_date)
        before = work["WHEN_DT"] < cutoff
        dropped = before.sum()
        work = work[~before].copy()
        if dropped > 0:
            st.info(f"🗑️ Removed {dropped:,} rows where Change Date < {cutoff.strftime('%d-%m-%Y')} · {len(work):,} rows remaining")

        # Build group key — ISO date string (handles mixed/unparseable gracefully)
        work["WHEN_KEY"] = work["WHEN_DT"].dt.strftime("%Y-%m-%d")
        bad_mask = work["WHEN_KEY"].isna()
        work.loc[bad_mask, "WHEN_KEY"] = work.loc[bad_mask, "WHEN_CHG"].astype(str).str.strip()

        # Created as int64 for fast max comparison (BUG FIX: use min int64 not Python overflow)
        work["CRT_INT"] = work["CREATED_DT"].values.astype("datetime64[ns]").astype("int64")
        work.loc[work["CREATED_DT"].isna(), "CRT_INT"] = -9223372036854775808

        # Clean CT values
        if auto_clean:
            work["CT_NUM"] = work["TOTAL_CT"].apply(clean_ctc)
            work["CT_FINAL"] = work["CT_NUM"].where(work["CT_NUM"].notna(), work["TOTAL_CT"])
        else:
            work["CT_FINAL"] = work["TOTAL_CT"]

        # DEDUPLICATE — fast: idxmax picks row with MAX CRT_INT per (EMP_ID+WHEN_KEY)
        idx = work.groupby(["EMP_ID","WHEN_KEY"])["CRT_INT"].idxmax()
        deduped = work.loc[idx].copy()

        # Pretty date label DD-MM-YYYY
        def pretty(k):
            try: return datetime.strptime(str(k), "%Y-%m-%d").strftime("%d-%m-%Y")
            except: return str(k)

        deduped["WHEN_PRETTY"] = deduped["WHEN_KEY"].apply(pretty)
        deduped = deduped.sort_values(["EMP_ID","WHEN_KEY"])

        # PIVOT — fast: pivot() is 6x faster than pivot_table() for deduplicated data
        pivot = deduped.pivot(
            index=["EMP_ID","EMP_NAME"],
            columns="WHEN_PRETTY",
            values="CT_FINAL"
        ).reset_index()
        pivot.columns.name = None

        # Sort date columns chronologically
        date_cols, other_cols = [], []
        for c in pivot.columns:
            if c in ["EMP_ID","EMP_NAME"]: continue
            try: datetime.strptime(str(c), "%d-%m-%Y"); date_cols.append(c)
            except: other_cols.append(c)

        date_cols_sorted = sorted(date_cols, key=lambda x: datetime.strptime(x, "%d-%m-%Y"))
        ordered = ["EMP_ID","EMP_NAME"] + date_cols_sorted + other_cols
        pivot = pivot[[c for c in ordered if c in pivot.columns]]

        # Rename back to original column names
        pivot = pivot.rename(columns={"EMP_ID": emp_col, "EMP_NAME": name_col})

        # Drop all-empty columns
        if drop_empty:
            pivot = pivot.dropna(axis=1, how="all")

        # Replace None/NaN with empty string for clean display
        pivot = pivot.fillna("")

        # Pre-generate download bytes INSIDE the spinner so they're always ready
        st.session_state["xlsx_bytes"] = to_excel_bytes(pivot)
        st.session_state["csv_bytes"]  = pivot.to_csv(index=False).encode("utf-8")

        # Store in session state
        st.session_state["pivot"]   = pivot
        st.session_state["deduped"] = deduped
        st.session_state["total"]   = len(df)

# ── Results ───────────────────────────────────────────────────────────────────
if "pivot" in st.session_state:
    pivot   = st.session_state["pivot"]
    deduped = st.session_state["deduped"]
    total   = st.session_state["total"]

    n_employees = pivot.shape[0]
    n_dates     = pivot.shape[1] - 2
    n_deduped   = len(deduped)

    st.divider()

    # Metric cards
    st.markdown(f"""
    <div class="metric-row">
      <div class="metric-card">
        <div class="metric-val">{total:,}</div>
        <div class="metric-label">Input Rows</div>
      </div>
      <div class="metric-card">
        <div class="metric-val">{n_deduped:,}</div>
        <div class="metric-label">After Dedup</div>
      </div>
      <div class="metric-card">
        <div class="metric-val">{n_employees:,}</div>
        <div class="metric-label">Employees</div>
      </div>
      <div class="metric-card">
        <div class="metric-val">{n_dates}</div>
        <div class="metric-label">Date Columns</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    st.success(f"✅ Done! {n_employees:,} employees · {n_dates} change-date columns · {pivot.shape[1]} total columns")

    # Preview
    st.markdown('<div class="section-head"><span>04</span> Preview (first 15 rows)</div>', unsafe_allow_html=True)
    st.dataframe(pivot.head(15), use_container_width=True, hide_index=True)

    # Downloads
    st.markdown('<div class="section-head"><span>05</span> Download</div>', unsafe_allow_html=True)
    dl1, dl2 = st.columns(2)

    with dl1:
        st.download_button(
            "⬇ Download Excel (.xlsx)",
            data=st.session_state["xlsx_bytes"],
            file_name="employee_ct_pivoted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="dl_xlsx"
        )
    with dl2:
        st.download_button(
            "⬇ Download CSV (.csv)",
            data=st.session_state["csv_bytes"],
            file_name="employee_ct_pivoted.csv",
            mime="text/csv",
            use_container_width=True,
            key="dl_csv"
        )

    # Diagnostics
    with st.expander("🔍 Diagnostics"):
        st.markdown("**Sample deduped data (long form)**")
        st.dataframe(
            deduped[["EMP_ID","EMP_NAME","WHEN_CHG","CREATED","CT_FINAL"]].head(20),
            use_container_width=True, hide_index=True
        )
        st.markdown("**Keys with multiple input rows (top 10)**")
        dup = deduped.groupby(["EMP_ID","WHEN_KEY"]).size().reset_index(name="count")
        dup = dup[dup["count"] > 1].sort_values("count", ascending=False)
        if len(dup) == 0:
            st.write("No duplicates found.")
        else:
            st.dataframe(dup.head(10), use_container_width=True, hide_index=True)
        st.markdown(f"**Output columns ({pivot.shape[1]}):** " + ", ".join(str(c) for c in list(pivot.columns)[:20]) + ("..." if pivot.shape[1]>20 else ""))
