import io
import re
import hmac
import hashlib
from datetime import datetime
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="CT Data Transformer",
    page_icon="🔄",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&display=swap');
html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
.stApp { background: #f7f5f2; }
.block-container { padding: 2.5rem 2rem 4rem 2rem; max-width: 1100px; }
#MainMenu, footer, header { visibility: hidden; }
.stDeployButton { display: none; }
.login-box { max-width: 420px; margin: 60px auto 0; padding: 44px 48px; background: #fff; border: 1px solid #e5e1db; border-radius: 10px; border-top: 3px solid #d4521a; box-shadow: 0 4px 24px rgba(0,0,0,0.07); }
.login-icon { font-size: 30px; margin-bottom: 8px; }
.login-title { font-size: 22px; font-weight: 700; color: #1a1714; margin-bottom: 4px; }
.login-sub { font-size: 13px; color: #8a7f75; margin-bottom: 28px; }
.admin-badge { display:inline-block; background:#d4521a; color:#fff; font-size:10px; font-weight:700; letter-spacing:.06em; text-transform:uppercase; padding:2px 8px; border-radius:3px; margin-left:6px; }
.hero { background: #1a1714; border-radius: 8px; padding: 36px 44px 32px; margin-bottom: 28px; position: relative; overflow: hidden; }
.hero::before { content: ''; position: absolute; top: -60px; right: -60px; width: 280px; height: 280px; background: radial-gradient(circle, rgba(212,82,26,0.25) 0%, transparent 70%); pointer-events: none; }
.hero-label { font-size: 11px; font-weight: 600; letter-spacing: .12em; text-transform: uppercase; color: #d4521a; display: flex; align-items: center; gap: 8px; margin-bottom: 12px; }
.hero-label::before { content: ''; width: 20px; height: 2px; background: #d4521a; }
.hero-title { font-size: 34px; font-weight: 300; color: #f5f0eb; letter-spacing: -.02em; margin-bottom: 6px; }
.hero-title strong { font-weight: 700; }
.hero-sub { font-size: 13px; color: #8a7f75; line-height: 1.6; max-width: 520px; }
.hero-user { position: absolute; top: 24px; right: 28px; font-size: 12px; color: #8a7f75; }
.hero-user strong { color: #d4521a; }
.metric-row { display: flex; gap: 12px; margin-bottom: 24px; }
.metric-card { flex: 1; background: #fff; border: 1px solid #e5e1db; border-radius: 6px; padding: 18px 20px; border-left: 3px solid #d4521a; }
.metric-val { font-size: 26px; font-weight: 700; font-family: 'DM Mono', monospace; color: #d4521a; line-height: 1; }
.metric-label { font-size: 10px; font-weight: 600; letter-spacing: .08em; text-transform: uppercase; color: #8a7f75; margin-top: 4px; }
.section-head { font-size: 10px; font-weight: 600; letter-spacing: .1em; text-transform: uppercase; color: #8a7f75; display: flex; align-items: center; gap: 8px; margin-bottom: 14px; padding-bottom: 10px; border-bottom: 1px solid #e5e1db; }
.section-head span { background: #d4521a; color: #fff; font-size: 9px; padding: 2px 7px; border-radius: 2px; }
[data-testid="stFileUploadDropzone"] { background: #fff !important; border: 2px dashed #d4c8bc !important; border-radius: 6px !important; }
[data-testid="stFileUploadDropzone"]:hover { border-color: #d4521a !important; }
[data-testid="stSelectbox"] > div > div { background: #fff !important; border: 1px solid #e5e1db !important; border-radius: 4px !important; font-size: 13px !important; }
[data-testid="stTextInput"] input { border: 1px solid #e5e1db !important; border-radius: 4px !important; font-size: 14px !important; }
[data-testid="stTextInput"] input:focus { border-color: #d4521a !important; box-shadow: none !important; }
.stButton > button, .stDownloadButton > button { border-radius: 4px !important; font-family: 'DM Sans', sans-serif !important; font-weight: 600 !important; font-size: 13px !important; transition: all .2s !important; }
.stButton > button[kind="primary"], .stDownloadButton > button { background: #d4521a !important; border: none !important; color: #fff !important; }
.stButton > button[kind="primary"]:hover, .stDownloadButton > button:hover { background: #c4420a !important; transform: translateY(-1px) !important; box-shadow: 0 4px 16px rgba(212,82,26,0.3) !important; }
.stSuccess { background: rgba(42,122,75,.07) !important; border-color: rgba(42,122,75,.25) !important; border-radius: 4px !important; }
.stInfo { background: rgba(212,82,26,.06) !important; border-color: rgba(212,82,26,.2) !important; border-radius: 4px !important; }
.stError, .stWarning { border-radius: 4px !important; }
[data-testid="stDataFrame"] { border: 1px solid #e5e1db !important; border-radius: 6px !important; overflow: hidden; }
hr { border-color: #e5e1db !important; }
[data-testid="stExpander"] { background: #fff !important; border: 1px solid #e5e1db !important; border-radius: 6px !important; }
</style>
""", unsafe_allow_html=True)

# ── AUTH ──────────────────────────────────────────────────────────────────────

ADMIN_EMAIL = "vivek.naiwal@cars24.com"

def get_users():
    """
    Load users from st.secrets.
    secrets.toml format:
        [users]
        "vivek.naiwal@cars24.com" = "Vivek@007"
        "payroll1@cars24.com"     = "Pass@123"
    """
    try:
        return {k.strip().lower(): v for k, v in dict(st.secrets["users"]).items()}
    except Exception:
        # Hardcoded fallback — only used locally, never in production
        return {
            "vivek.naiwal@cars24.com": "Vivek@007"
        }

def check_login(email, password):
    users = get_users()
    email = email.strip().lower()
    if email not in users:
        return False
    return hmac.compare_digest(
        hashlib.sha256(password.encode()).hexdigest(),
        hashlib.sha256(str(users[email]).encode()).hexdigest()
    )

def show_login():
    _, center, _ = st.columns([1, 1.4, 1])
    with center:
        st.markdown("""
        <div class="login-box">
            <div class="login-icon">🔐</div>
            <div class="login-title">Payroll Portal</div>
            <div class="login-sub">CT Data Transformer — Restricted Access<br>Cars24 Internal Use Only</div>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        email    = st.text_input("Email", placeholder="you@cars24.com", key="li_email")
        password = st.text_input("Password", type="password", placeholder="Enter password", key="li_pass")
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("Sign In →", type="primary", use_container_width=True):
            if not email or not password:
                st.error("Please enter both email and password.")
            elif check_login(email, password):
                st.session_state["auth"]       = True
                st.session_state["user_email"] = email.strip().lower()
                st.session_state["is_admin"]   = (email.strip().lower() == ADMIN_EMAIL.lower())
                st.rerun()
            else:
                st.error("❌ Wrong email or password. Contact vivek.naiwal@cars24.com")
        st.markdown("<div style='text-align:center;font-size:11px;color:#b0a89f;margin-top:16px;'>Authorised personnel only</div>", unsafe_allow_html=True)

if not st.session_state.get("auth"):
    show_login()
    st.stop()

user_email = st.session_state["user_email"]
is_admin   = st.session_state["is_admin"]

# ── COLUMN CANDIDATES ─────────────────────────────────────────────────────────

COLUMN_CANDIDATES = {
    "emp_id":       ["emp id", "empid", "employee id", "ecode", "emp_id", "employee_code"],
    "emp_name":     ["employee name", "name", "emp name", "employee"],
    "when_change":  ["when was the change", "when_was_the_change", "effective date", "when changed", "change date", "whenchange"],
    "total_ctc":    ["total ct", "total ctc", "ctc", "total_ct", "total"],
    "created_date": ["created date", "created_date", "created on", "created", "created_at", "createdat"]
}

# ── HELPERS ───────────────────────────────────────────────────────────────────

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
    try: return list(columns).index(value) + 1
    except: return 0

# ── HERO ──────────────────────────────────────────────────────────────────────

admin_badge = '<span class="admin-badge">Admin</span>' if is_admin else ""
st.markdown(f"""
<div class="hero">
  <div class="hero-user">{'👑 ' if is_admin else ''}<strong>{user_email}</strong>{admin_badge}</div>
  <div class="hero-label">Payroll Internal Tool</div>
  <div class="hero-title">CT Data <strong>Transformer</strong></div>
  <div class="hero-sub">Upload your employee CT file. Keeps the latest record per change date
  and pivots — one row per employee, each change date as a column.</div>
</div>
""", unsafe_allow_html=True)

_, signout_col = st.columns([6, 1])
with signout_col:
    if st.button("Sign Out", key="logout"):
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.rerun()

# Admin panel — only vivek sees this
if is_admin:
    with st.expander("⚙️ Admin — Manage Users", expanded=False):
        st.markdown("**Add/remove users in Streamlit Secrets** → [users] section:")
        st.code("""[users]
"vivek.naiwal@cars24.com" = "Vivek@007"
"payroll1@cars24.com"     = "Password1"
"payroll2@cars24.com"     = "Password2"
"manager@cars24.com"      = "Password3"
""", language="toml")
        st.info("Go to Streamlit Cloud → your app → ⋮ → Settings → Secrets to add/remove users. Changes apply immediately on next login.")
        st.markdown("**Currently signed-in user:** " + user_email + "")

st.divider()

# ── STEP 1: UPLOAD ────────────────────────────────────────────────────────────

st.markdown('<div class="section-head"><span>01</span> Upload File</div>', unsafe_allow_html=True)

uploaded = st.file_uploader("Drop your CSV or Excel file here",
    type=["csv", "xlsx", "xls"], accept_multiple_files=False, label_visibility="collapsed")

if uploaded is None:
    st.info("📂 Upload a file above to get started.")
    st.stop()

try:
    raw_bytes = uploaded.read()
    if uploaded.name.lower().endswith(".csv"):
        df = pd.read_csv(io.BytesIO(raw_bytes), dtype=str)
    else:
        df = pd.read_excel(io.BytesIO(raw_bytes), dtype=str)
except Exception as e:
    st.error(f"❌ Failed to read file: {e}")
    st.stop()

st.success(f"✅ {uploaded.name} loaded — {len(df):,} rows · {len(df.columns)} columns")
st.divider()

# ── STEP 2: CONFIGURE ─────────────────────────────────────────────────────────

st.markdown('<div class="section-head"><span>02</span> Configure</div>', unsafe_allow_html=True)

col_opt1, col_opt2, col_opt3 = st.columns(3)
with col_opt1: dayfirst   = st.checkbox("Dates are DD-MM-YYYY (day first)", value=True)
with col_opt2: auto_clean = st.checkbox("Clean Total CT (remove commas/currency)", value=True)
with col_opt3: drop_empty = st.checkbox("Drop all-empty columns in output", value=True)

fc1, fc2 = st.columns([1, 3])
with fc1:
    filter_date = st.date_input("Remove Change Dates before",
        value=pd.Timestamp("2024-01-01"),
        min_value=pd.Timestamp("2000-01-01"),
        max_value=pd.Timestamp("2030-12-31"),
        help="Rows with Change Date before this are excluded")

st.markdown("**Column Mapping** — auto-detected from your headers:")
detected = {k: find_column(df.columns.tolist(), v) for k, v in COLUMN_CANDIDATES.items()}
options_with_none = [None] + list(df.columns)

mc1, mc2, mc3, mc4, mc5 = st.columns(5)
with mc1: emp_col     = st.selectbox("EMP ID",           options_with_none, index=safe_index(df.columns, detected["emp_id"]))
with mc2: name_col    = st.selectbox("Employee Name",    options_with_none, index=safe_index(df.columns, detected["emp_name"]))
with mc3: when_col    = st.selectbox("When Was Change?", options_with_none, index=safe_index(df.columns, detected["when_change"]))
with mc4: ct_col      = st.selectbox("Total CT",         options_with_none, index=safe_index(df.columns, detected["total_ctc"]))
with mc5: created_col = st.selectbox("Created Date",     options_with_none, index=safe_index(df.columns, detected["created_date"]))

missing = [k for k, v in {"EMP ID": emp_col, "Employee Name": name_col,
    "When Was Change?": when_col, "Total CT": ct_col, "Created Date": created_col}.items() if not v]
if missing:
    st.error(f"❌ Please map these columns: {', '.join(missing)}")
    st.stop()

st.divider()

# ── STEP 3: PROCESS ───────────────────────────────────────────────────────────

st.markdown('<div class="section-head"><span>03</span> Process</div>', unsafe_allow_html=True)

if st.button("⚡  Process & Generate Output", type="primary", use_container_width=True):
    with st.spinner("Processing..."):
        work = df[[emp_col, name_col, when_col, ct_col, created_col]].copy()
        work.columns = ["EMP_ID", "EMP_NAME", "WHEN_CHG", "TOTAL_CT", "CREATED"]

        work["WHEN_DT"]    = try_parse_dates(work["WHEN_CHG"], dayfirst=dayfirst)
        work["CREATED_DT"] = try_parse_dates(work["CREATED"],  dayfirst=dayfirst)

        cutoff = pd.Timestamp(filter_date)
        before = work["WHEN_DT"] < cutoff
        dropped = before.sum()
        work = work[~before].copy()
        if dropped > 0:
            st.info(f"🗑️ Removed {dropped:,} rows before {cutoff.strftime('%d-%m-%Y')} · {len(work):,} rows remaining")

        work["WHEN_KEY"] = work["WHEN_DT"].dt.strftime("%Y-%m-%d")
        bad_mask = work["WHEN_KEY"].isna()
        work.loc[bad_mask, "WHEN_KEY"] = work.loc[bad_mask, "WHEN_CHG"].astype(str).str.strip()

        work["CRT_INT"] = work["CREATED_DT"].values.astype("datetime64[ns]").astype("int64")
        work.loc[work["CREATED_DT"].isna(), "CRT_INT"] = -9223372036854775808

        if auto_clean:
            work["CT_NUM"]   = work["TOTAL_CT"].apply(clean_ctc)
            work["CT_FINAL"] = work["CT_NUM"].where(work["CT_NUM"].notna(), work["TOTAL_CT"])
        else:
            work["CT_FINAL"] = work["TOTAL_CT"]

        idx     = work.groupby(["EMP_ID", "WHEN_KEY"])["CRT_INT"].idxmax()
        deduped = work.loc[idx].copy()

        def pretty(k):
            try: return datetime.strptime(str(k), "%Y-%m-%d").strftime("%d-%m-%Y")
            except: return str(k)

        deduped["WHEN_PRETTY"] = deduped["WHEN_KEY"].apply(pretty)
        deduped = deduped.sort_values(["EMP_ID", "WHEN_KEY"])

        pivot = deduped.pivot(index=["EMP_ID", "EMP_NAME"], columns="WHEN_PRETTY", values="CT_FINAL").reset_index()
        pivot.columns.name = None

        date_cols, other_cols = [], []
        for c in pivot.columns:
            if c in ["EMP_ID", "EMP_NAME"]: continue
            try: datetime.strptime(str(c), "%d-%m-%Y"); date_cols.append(c)
            except: other_cols.append(c)

        date_cols = sorted(date_cols, key=lambda x: datetime.strptime(x, "%d-%m-%Y"))
        pivot = pivot[["EMP_ID", "EMP_NAME"] + date_cols + other_cols]
        pivot = pivot.rename(columns={"EMP_ID": emp_col, "EMP_NAME": name_col})

        if drop_empty:
            pivot = pivot.dropna(axis=1, how="all")
        pivot = pivot.fillna("")

        st.session_state["xlsx_bytes"] = to_excel_bytes(pivot)
        st.session_state["csv_bytes"]  = pivot.to_csv(index=False).encode("utf-8")
        st.session_state["pivot"]      = pivot
        st.session_state["deduped"]    = deduped
        st.session_state["total"]      = len(df)

# ── RESULTS ───────────────────────────────────────────────────────────────────

if "pivot" in st.session_state:
    pivot   = st.session_state["pivot"]
    deduped = st.session_state["deduped"]
    total   = st.session_state["total"]

    n_emp   = pivot.shape[0]
    n_dates = pivot.shape[1] - 2
    n_dedup = len(deduped)

    st.divider()
    st.markdown(f"""
    <div class="metric-row">
      <div class="metric-card"><div class="metric-val">{total:,}</div><div class="metric-label">Input Rows</div></div>
      <div class="metric-card"><div class="metric-val">{n_dedup:,}</div><div class="metric-label">After Dedup</div></div>
      <div class="metric-card"><div class="metric-val">{n_emp:,}</div><div class="metric-label">Employees</div></div>
      <div class="metric-card"><div class="metric-val">{n_dates}</div><div class="metric-label">Date Columns</div></div>
    </div>
    """, unsafe_allow_html=True)

    st.success(f"✅ Done! {n_emp:,} employees · {n_dates} change-date columns · {pivot.shape[1]} total columns")

    st.markdown('<div class="section-head"><span>04</span> Preview (first 15 rows)</div>', unsafe_allow_html=True)
    st.dataframe(pivot.head(15), use_container_width=True, hide_index=True)

    st.markdown('<div class="section-head"><span>05</span> Download</div>', unsafe_allow_html=True)
    dl1, dl2 = st.columns(2)
    with dl1:
        st.download_button("⬇ Download Excel (.xlsx)",
            data=st.session_state["xlsx_bytes"], file_name="employee_ct_pivoted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True, key="dl_xlsx")
    with dl2:
        st.download_button("⬇ Download CSV (.csv)",
            data=st.session_state["csv_bytes"], file_name="employee_ct_pivoted.csv",
            mime="text/csv", use_container_width=True, key="dl_csv")

    with st.expander("🔍 Diagnostics"):
        st.dataframe(deduped[["EMP_ID", "EMP_NAME", "WHEN_CHG", "CREATED", "CT_FINAL"]].head(20),
            use_container_width=True, hide_index=True)
        dup = deduped.groupby(["EMP_ID", "WHEN_KEY"]).size().reset_index(name="count")
        dup = dup[dup["count"] > 1].sort_values("count", ascending=False)
        st.write("Duplicate keys (top 10):" if len(dup) else "No duplicates found.")
        if len(dup):
            st.dataframe(dup.head(10), use_container_width=True, hide_index=True)
