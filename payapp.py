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
    try:
        return {k.strip().lower(): v for k, v in dict(st.secrets["users"]).items()}
    except Exception:
        return {"vivek.naiwal@cars24.com": "Vivek@007"}

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
    "when_change":  ["when was the change", "when was the change?", "when_was_the_change", "effective date", "when changed", "change date", "whenchange"],
    "total_ctc":    ["total ctc", "total ct", "ctc", "total_ctc", "total_ct", "total"],
    "fixed_ctc":    ["fixed ctc", "fixed_ctc", "fixed ct", "fixed salary", "fixed"],
    "variable_pay": ["total variable pay", "variable pay", "total_variable_pay", "variable", "var pay"],
    "perf_var":     ["performance linked variable", "perf linked variable", "perf var", "performance variable"],
    "ret_bonus":    ["retention bonus", "ret bonus", "retention"],
    "perf_inc":     ["performance linked incentive", "perf linked incentive", "perf incentive", "performance incentive"],
    "total_pay":    ["total pay", "total_pay", "totalpay"],
    "status":       ["status"],
    "created_date": ["created date", "created_date", "created on", "created", "created_at", "createdat"],
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

def clean_numeric(x):
    if pd.isna(x): return pd.NA
    if isinstance(x, (int, float)): return x
    s = re.sub(r"[^\d\.\-]", "", str(x).strip())
    try: return float(s)
    except: return pd.NA

def safe_index(columns, value):
    if value is None: return 0
    try: return list(columns).index(value) + 1
    except: return 0


# ── CORE TRANSFORM: 1 row per employee, Latest→2nd→3rd records side-by-side ──

def build_employee_records(df, col_map, dayfirst, auto_clean, date_cutoff=None):
    """
    Logic mirrors the VBA macro exactly:
    1. For each employee, gather all records.
    2. Optionally filter rows where Effective Date < date_cutoff.
    3. If any non-Discarded records exist, exclude Discarded ones.
    4. Sort: Active first, then Archived, then by Created Date DESC.
    5. Take top-3 records → Latest, 2nd Latest, 3rd Latest.
    6. Output 1 row per employee.
    """
    # Reset index — critical when df arrived with a non-default index
    # (e.g. after skiprows detection); .values assignment would silently
    # misalign otherwise.
    df = df.reset_index(drop=True)

    # Build work df column-by-column to handle duplicate src cols safely.
    work = pd.DataFrame(index=df.index)
    for key, src_col in col_map.items():
        # Handle duplicate column names: df[col] returns DataFrame if duplicated
        col_idx = list(df.columns).index(src_col)
        work[key] = df.iloc[:, col_idx].values

    work["WHEN_DT"]    = try_parse_dates(work["when_change"], dayfirst=dayfirst)
    work["CREATED_DT"] = try_parse_dates(work["created_date"], dayfirst=dayfirst)

    # Apply date cutoff filter if provided
    if date_cutoff is not None:
        cutoff_ts = pd.Timestamp(date_cutoff)
        before    = work["WHEN_DT"] < cutoff_ts
        dropped   = int(before.sum())
        work      = work[~before].copy().reset_index(drop=True)
        if dropped > 0:
            import streamlit as _st
            _st.info(f"🗑️ Removed {dropped:,} rows with Effective Date before {cutoff_ts.strftime('%d-%b-%Y')} · {len(work):,} rows remaining")

    num_cols = ["fixed_ctc", "total_pay", "variable_pay", "total_ctc", "perf_var", "ret_bonus", "perf_inc"]
    for c in num_cols:
        if auto_clean:
            work[c + "_N"] = work[c].apply(clean_numeric)
        else:
            work[c + "_N"] = pd.to_numeric(work[c], errors="coerce")

    # Status priority: Active=0, Archived=1, Discarded=2
    def status_rank(s):
        s = str(s).lower().strip()
        if s == "active":    return 0
        if s == "archived":  return 1
        return 2  # discarded or unknown

    work["STATUS_RANK"] = work["status"].apply(status_rank)

    rows = []
    emp_col = "emp_id"

    for emp_id, grp in work.groupby(emp_col, sort=False):
        # Filter out Discarded if any non-Discarded exist
        non_disc = grp[grp["STATUS_RANK"] < 2]
        pool = non_disc if len(non_disc) > 0 else grp

        # Sort: Status rank ASC, then Created Date DESC
        pool = pool.copy()
        pool["CRT_TS"] = pool["CREATED_DT"].values.astype("datetime64[ns]").astype("int64")
        pool = pool.sort_values(["STATUS_RANK", "CRT_TS"], ascending=[True, False])

        top3 = pool.head(3).reset_index(drop=True)

        row = {
            "EMP ID":         emp_id,
            "Employee Name":  top3.loc[0, "emp_name"] if len(top3) > 0 else "",
        }

        labels = ["LATEST / ACTIVE RECORD", "2nd LATEST RECORD", "3rd LATEST RECORD"]
        fields = [
            ("Effective Date",  "WHEN_DT"),
            ("Fixed CTC",       "fixed_ctc_N"),
            ("Total Pay",       "total_pay_N"),
            ("Total Var. Pay",  "variable_pay_N"),
            ("Total CTC",       "total_ctc_N"),
            ("Perf. Linked Var.","perf_var_N"),
            ("Retention Bonus", "ret_bonus_N"),
            ("Perf. Linked Inc.","perf_inc_N"),
        ]

        for i, label in enumerate(labels):
            if i < len(top3):
                r = top3.iloc[i]
                for fname, fkey in fields:
                    col_name = f"{label} | {fname}"
                    val = r[fkey]
                    if fname == "Effective Date":
                        row[col_name] = val.strftime("%d-%b-%Y") if pd.notna(val) else ""
                    else:
                        row[col_name] = int(val) if pd.notna(val) and val == val else ""
            else:
                for fname, _ in fields:
                    row[f"{label} | {fname}"] = ""

        rows.append(row)

    return pd.DataFrame(rows)


# ── EXCEL BUILDER — matches Output sheet layout exactly ──────────────────────

def to_excel_bytes_output(result_df):
    """
    Mirrors the Output sheet structure:
    Row 1: Title banner
    Row 2: Group headers  → Employee Info | LATEST/ACTIVE | separator | 2nd LATEST | separator | 3rd LATEST
    Row 3: Sub-headers    → S.NO | EMP ID | Name | Eff Date | Fixed | TotalPay | VarPay | TotCTC | PerfVar | RetBonus | PerfInc | [sep] | ...×2
    Row 4+: Data rows with alternating blue/orange/green colour bands
    """
    import xlsxwriter

    buf = io.BytesIO()
    wb  = xlsxwriter.Workbook(buf, {"in_memory": True, "constant_memory": False})
    ws  = wb.add_worksheet("Output")

    # ── Formats ──────────────────────────────────────────────────────────────
    # Title row
    fmt_title = wb.add_format({
        "bold": True, "font_size": 12, "font_name": "Arial",
        "bg_color": "#1A1714", "font_color": "#F5F0EB",
        "align": "left", "valign": "vcenter", "border": 1
    })
    # Group header — Employee Info (blue-ish)
    fmt_grp_emp = wb.add_format({
        "bold": True, "font_size": 10, "font_name": "Arial",
        "bg_color": "#BDD7EE", "font_color": "#1A1714",
        "align": "center", "valign": "vcenter", "border": 1
    })
    # Group header — Latest (blue tones)
    fmt_grp_latest = wb.add_format({
        "bold": True, "font_size": 10, "font_name": "Arial",
        "bg_color": "#2E75B6", "font_color": "#FFFFFF",
        "align": "center", "valign": "vcenter", "border": 1
    })
    # Group header — 2nd Latest (orange tones)
    fmt_grp_2nd = wb.add_format({
        "bold": True, "font_size": 10, "font_name": "Arial",
        "bg_color": "#D4521A", "font_color": "#FFFFFF",
        "align": "center", "valign": "vcenter", "border": 1
    })
    # Group header — 3rd Latest (green tones)
    fmt_grp_3rd = wb.add_format({
        "bold": True, "font_size": 10, "font_name": "Arial",
        "bg_color": "#375623", "font_color": "#FFFFFF",
        "align": "center", "valign": "vcenter", "border": 1
    })
    # Separator column
    fmt_sep = wb.add_format({
        "bg_color": "#F2F2F2", "border": 1
    })
    # Sub-header base
    def sub_hdr(bg, fc="#1A1714"):
        return wb.add_format({
            "bold": True, "font_size": 9, "font_name": "Arial",
            "bg_color": bg, "font_color": fc,
            "align": "center", "valign": "vcenter",
            "text_wrap": True, "border": 1
        })
    fmt_sub_emp   = sub_hdr("#BDD7EE")
    fmt_sub_sno   = sub_hdr("#BDD7EE")
    fmt_sub_lat   = sub_hdr("#DEEAF1")
    fmt_sub_2nd   = sub_hdr("#FCE4D6")
    fmt_sub_3rd   = sub_hdr("#E2EFDA")

    # Data formats
    def data_fmt(bg, num=False):
        d = {
            "font_size": 10, "font_name": "Arial",
            "bg_color": bg, "align": "center", "valign": "vcenter", "border": 1,
            "border_color": "#B4B4B4"
        }
        if num:
            d["num_format"] = "#,##0"
            d["align"] = "right"
        return wb.add_format(d)

    # Alternating fills per section
    # Latest: blue
    fmt_lat_even_txt = data_fmt("#DEEAF1")
    fmt_lat_odd_txt  = data_fmt("#EBF3F9")
    fmt_lat_even_num = data_fmt("#DEEAF1", num=True)
    fmt_lat_odd_num  = data_fmt("#EBF3F9", num=True)
    # 2nd: orange
    fmt_2nd_even_txt = data_fmt("#FCE4D6")
    fmt_2nd_odd_txt  = data_fmt("#FFF0E8")
    fmt_2nd_even_num = data_fmt("#FCE4D6", num=True)
    fmt_2nd_odd_num  = data_fmt("#FFF0E8", num=True)
    # 3rd: green
    fmt_3rd_even_txt = data_fmt("#E2EFDA")
    fmt_3rd_odd_txt  = data_fmt("#F0F7EC")
    fmt_3rd_even_num = data_fmt("#E2EFDA", num=True)
    fmt_3rd_odd_num  = data_fmt("#F0F7EC", num=True)
    # EMP cols: solid blue
    fmt_emp_even = data_fmt("#BDD7EE")
    fmt_emp_odd  = data_fmt("#DEEAF1")
    fmt_emp_even_bold = wb.add_format({"bold": True, "font_size": 10, "font_name": "Arial", "bg_color": "#BDD7EE", "align": "center", "valign": "vcenter", "border": 1, "border_color": "#B4B4B4"})
    fmt_emp_odd_bold  = wb.add_format({"bold": True, "font_size": 10, "font_name": "Arial", "bg_color": "#DEEAF1", "align": "center", "valign": "vcenter", "border": 1, "border_color": "#B4B4B4"})
    fmt_sno_even = wb.add_format({"bold": True, "font_size": 10, "font_name": "Arial", "bg_color": "#BDD7EE", "align": "center", "valign": "vcenter", "border": 1})
    fmt_sno_odd  = wb.add_format({"bold": True, "font_size": 10, "font_name": "Arial", "bg_color": "#DEEAF1", "align": "center", "valign": "vcenter", "border": 1})
    fmt_name_even = wb.add_format({"font_size": 10, "font_name": "Arial", "bg_color": "#BDD7EE", "align": "left", "valign": "vcenter", "border": 1})
    fmt_name_odd  = wb.add_format({"font_size": 10, "font_name": "Arial", "bg_color": "#DEEAF1", "align": "left", "valign": "vcenter", "border": 1})
    fmt_date_even = wb.add_format({"font_size": 10, "font_name": "Arial", "bg_color": "#DEEAF1", "align": "center", "valign": "vcenter", "border": 1, "border_color": "#B4B4B4"})
    fmt_date_odd  = wb.add_format({"font_size": 10, "font_name": "Arial", "bg_color": "#EBF3F9", "align": "center", "valign": "vcenter", "border": 1, "border_color": "#B4B4B4"})
    fmt_date_2nd_even = wb.add_format({"font_size": 10, "font_name": "Arial", "bg_color": "#FCE4D6", "align": "center", "valign": "vcenter", "border": 1, "border_color": "#B4B4B4"})
    fmt_date_2nd_odd  = wb.add_format({"font_size": 10, "font_name": "Arial", "bg_color": "#FFF0E8", "align": "center", "valign": "vcenter", "border": 1, "border_color": "#B4B4B4"})
    fmt_date_3rd_even = wb.add_format({"font_size": 10, "font_name": "Arial", "bg_color": "#E2EFDA", "align": "center", "valign": "vcenter", "border": 1, "border_color": "#B4B4B4"})
    fmt_date_3rd_odd  = wb.add_format({"font_size": 10, "font_name": "Arial", "bg_color": "#F0F7EC", "align": "center", "valign": "vcenter", "border": 1, "border_color": "#B4B4B4"})

    # ── Column widths ─────────────────────────────────────────────────────────
    ws.set_column(0,  0,  5)   # S.NO
    ws.set_column(1,  1,  10)  # EMP ID
    ws.set_column(2,  2,  22)  # Employee Name
    ws.set_column(3,  3,  14)  # Eff Date (Latest)
    ws.set_column(4,  10, 14)  # Latest numeric cols
    ws.set_column(11, 11, 3)   # separator
    ws.set_column(12, 12, 14)  # Eff Date (2nd)
    ws.set_column(13, 19, 14)  # 2nd numeric cols
    ws.set_column(20, 20, 3)   # separator
    ws.set_column(21, 21, 14)  # Eff Date (3rd)
    ws.set_column(22, 28, 14)  # 3rd numeric cols

    # ── ROW 0: Title banner ───────────────────────────────────────────────────
    ws.set_row(0, 20)
    ws.merge_range(0, 0, 0, 28,
        "CTC TRANSFORMATION OUTPUT  —  1 Row per Employee  |  Latest → 2nd Latest → 3rd Latest",
        fmt_title)

    # ── ROW 1: Group headers ──────────────────────────────────────────────────
    ws.set_row(1, 20)
    ws.merge_range(1, 0, 1, 2,  "Employee Info",          fmt_grp_emp)
    ws.merge_range(1, 3, 1, 10, "LATEST / ACTIVE RECORD", fmt_grp_latest)
    ws.write(1, 11, "", fmt_sep)
    ws.merge_range(1, 12, 1, 19, "2nd LATEST RECORD",     fmt_grp_2nd)
    ws.write(1, 20, "", fmt_sep)
    ws.merge_range(1, 21, 1, 28, "3rd LATEST RECORD",     fmt_grp_3rd)

    # ── ROW 2: Sub-headers ────────────────────────────────────────────────────
    ws.set_row(2, 30)
    sub_labels = ["Effective Date", "Fixed CTC", "Total Pay", "Total Var. Pay",
                  "Total CTC", "Perf. Linked Var.", "Retention Bonus", "Perf. Linked Inc."]

    ws.write(2, 0,  "S.NO",          fmt_sub_sno)
    ws.write(2, 1,  "EMP ID",        fmt_sub_emp)
    ws.write(2, 2,  "Employee Name", fmt_sub_emp)
    for j, lbl in enumerate(sub_labels):
        ws.write(2, 3  + j, lbl, fmt_sub_lat)
    ws.write(2, 11, "", fmt_sep)
    for j, lbl in enumerate(sub_labels):
        ws.write(2, 12 + j, lbl, fmt_sub_2nd)
    ws.write(2, 20, "", fmt_sep)
    for j, lbl in enumerate(sub_labels):
        ws.write(2, 21 + j, lbl, fmt_sub_3rd)

    ws.freeze_panes(3, 3)

    # ── DATA ROWS ─────────────────────────────────────────────────────────────
    labels = ["LATEST / ACTIVE RECORD", "2nd LATEST RECORD", "3rd LATEST RECORD"]
    field_names = ["Effective Date", "Fixed CTC", "Total Pay", "Total Var. Pay",
                   "Total CTC", "Perf. Linked Var.", "Retention Bonus", "Perf. Linked Inc."]
    start_cols = [3, 12, 21]
    sep_cols   = [11, 20]

    for ri, row_data in result_df.iterrows():
        xl_row = ri + 3
        even   = (ri % 2 == 0)
        ws.set_row(xl_row, 18)

        # EMP info
        ws.write(xl_row, 0, ri + 1,                  fmt_sno_even  if even else fmt_sno_odd)
        ws.write(xl_row, 1, str(row_data["EMP ID"]), fmt_emp_even_bold if even else fmt_emp_odd_bold)
        ws.write(xl_row, 2, str(row_data["Employee Name"]), fmt_name_even if even else fmt_name_odd)

        # Separators
        for sc in sep_cols:
            ws.write(xl_row, sc, "", fmt_sep)

        # 3 record groups
        for gi, label in enumerate(labels):
            sc = start_cols[gi]
            # Date
            date_val = row_data.get(f"{label} | Effective Date", "")
            if gi == 0:
                date_fmt = fmt_date_even if even else fmt_date_odd
                num_fmt_txt = fmt_lat_even_txt  if even else fmt_lat_odd_txt
                num_fmt_num = fmt_lat_even_num  if even else fmt_lat_odd_num
            elif gi == 1:
                date_fmt = fmt_date_2nd_even if even else fmt_date_2nd_odd
                num_fmt_txt = fmt_2nd_even_txt  if even else fmt_2nd_odd_txt
                num_fmt_num = fmt_2nd_even_num  if even else fmt_2nd_odd_num
            else:
                date_fmt = fmt_date_3rd_even if even else fmt_date_3rd_odd
                num_fmt_txt = fmt_3rd_even_txt  if even else fmt_3rd_odd_txt
                num_fmt_num = fmt_3rd_even_num  if even else fmt_3rd_odd_num

            ws.write(xl_row, sc, date_val, date_fmt)

            # Numeric fields
            for fi, fname in enumerate(field_names[1:], start=1):
                val = row_data.get(f"{label} | {fname}", "")
                if val == "" or val is None:
                    ws.write(xl_row, sc + fi, "", num_fmt_txt)
                else:
                    try:
                        ws.write_number(xl_row, sc + fi, float(val), num_fmt_num)
                    except Exception:
                        ws.write(xl_row, sc + fi, val, num_fmt_txt)

    wb.close()
    return buf.getvalue()


# ── HERO ──────────────────────────────────────────────────────────────────────

admin_badge = '<span class="admin-badge">Admin</span>' if is_admin else ""
st.markdown(f"""
<div class="hero">
  <div class="hero-user">{'👑 ' if is_admin else ''}<strong>{user_email}</strong>{admin_badge}</div>
  <div class="hero-label">Payroll Internal Tool</div>
  <div class="hero-title">CT Data <strong>Transformer</strong></div>
  <div class="hero-sub">Upload your employee CT file. Outputs 1 row per employee with Latest, 2nd Latest,
  and 3rd Latest CTC records side-by-side — matching the Output sheet format.</div>
</div>
""", unsafe_allow_html=True)

_, signout_col = st.columns([6, 1])
with signout_col:
    if st.button("Sign Out", key="logout"):
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.rerun()

if is_admin:
    with st.expander("⚙️ Admin — Manage Users", expanded=False):
        st.markdown("**Add/remove users in Streamlit Secrets** → [users] section:")
        st.code("""[users]
"vivek.naiwal@cars24.com" = "Vivek@007"
"payroll1@cars24.com"     = "Password1"
"payroll2@cars24.com"     = "Password2"
"manager@cars24.com"      = "Password3"
""", language="toml")
        st.info("Go to Streamlit Cloud → your app → ⋮ → Settings → Secrets to add/remove users.")
        st.markdown("**Currently signed-in user:** " + user_email)

st.divider()

# ── STEP 1: UPLOAD ────────────────────────────────────────────────────────────

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

try:
    raw_bytes = uploaded.read()
    if uploaded.name.lower().endswith(".csv"):
        df = pd.read_csv(io.BytesIO(raw_bytes), dtype=str)
    else:
        # Try to find header row (skip rows until we see EMP ID-like header)
        for skip in range(0, 5):
            df = pd.read_excel(io.BytesIO(raw_bytes), dtype=str, skiprows=skip)
            df.columns = [c.strip() for c in df.columns]
            if any("emp" in c.lower() or "employee" in c.lower() for c in df.columns):
                break
except Exception as e:
    st.error(f"❌ Failed to read file: {e}")
    st.stop()

df.columns = [c.strip() for c in df.columns]
st.success(f"✅ {uploaded.name} loaded — {len(df):,} rows · {len(df.columns)} columns")
st.divider()

# ── STEP 2: CONFIGURE ─────────────────────────────────────────────────────────

st.markdown('<div class="section-head"><span>02</span> Configure</div>', unsafe_allow_html=True)

col_opt1, col_opt2 = st.columns(2)
with col_opt1: dayfirst   = st.checkbox("Dates are DD-MM-YYYY (day first)", value=True)
with col_opt2: auto_clean = st.checkbox("Clean numeric fields (remove commas/currency)", value=True)

fc1, _ = st.columns([1, 3])
with fc1:
    filter_date = st.date_input(
        "Remove Change Dates before",
        value=pd.Timestamp("2024-01-01"),
        min_value=pd.Timestamp("2000-01-01"),
        max_value=pd.Timestamp("2030-12-31"),
        help="Rows with Effective Date before this date are excluded from all records"
    )

detected = {k: find_column(df.columns.tolist(), v) for k, v in COLUMN_CANDIDATES.items()}
options_with_none = [None] + list(df.columns)

st.markdown("**Column Mapping** — auto-detected from your headers:")
mc1, mc2, mc3 = st.columns(3)
mc4, mc5, mc6 = st.columns(3)
mc7, mc8, mc9 = st.columns(3)
mc10, mc11, mc12 = st.columns(3)

with mc1:  emp_col    = st.selectbox("EMP ID",                    options_with_none, index=safe_index(df.columns, detected["emp_id"]))
with mc2:  name_col   = st.selectbox("Employee Name",             options_with_none, index=safe_index(df.columns, detected["emp_name"]))
with mc3:  when_col   = st.selectbox("Effective Date (When Change?)", options_with_none, index=safe_index(df.columns, detected["when_change"]))
with mc4:  fixed_col  = st.selectbox("Fixed CTC",                 options_with_none, index=safe_index(df.columns, detected["fixed_ctc"]))
with mc5:  totpay_col = st.selectbox("Total Pay",                 options_with_none, index=safe_index(df.columns, detected["total_pay"]))
with mc6:  var_col    = st.selectbox("Total Variable Pay",        options_with_none, index=safe_index(df.columns, detected["variable_pay"]))
with mc7:  total_col  = st.selectbox("Total CTC",                 options_with_none, index=safe_index(df.columns, detected["total_ctc"]))
with mc8:  perf_col   = st.selectbox("Perf. Linked Variable",     options_with_none, index=safe_index(df.columns, detected["perf_var"]))
with mc9:  ret_col    = st.selectbox("Retention Bonus",           options_with_none, index=safe_index(df.columns, detected["ret_bonus"]))
with mc10: perfinc_col= st.selectbox("Perf. Linked Incentive",   options_with_none, index=safe_index(df.columns, detected["perf_inc"]))
with mc11: status_col = st.selectbox("Status",                    options_with_none, index=safe_index(df.columns, detected["status"]))
with mc12: created_col= st.selectbox("Created Date",              options_with_none, index=safe_index(df.columns, detected["created_date"]))

required = {
    "EMP ID": emp_col, "Employee Name": name_col,
    "Effective Date": when_col, "Fixed CTC": fixed_col,
    "Total Pay": totpay_col, "Total Variable Pay": var_col,
    "Total CTC": total_col, "Status": status_col, "Created Date": created_col
}
missing = [k for k, v in required.items() if not v]
if missing:
    st.error(f"❌ Please map these required columns: {', '.join(missing)}")
    st.stop()

# Optional cols default to a zero-fill column if unmapped
def coalesce_col(col_name, df):
    if col_name:
        return col_name
    df["__ZERO__"] = "0"
    return "__ZERO__"

perf_col_r   = coalesce_col(perf_col,    df)
ret_col_r    = coalesce_col(ret_col,     df)
perfinc_col_r= coalesce_col(perfinc_col, df)

st.divider()

# ── STEP 3: PROCESS ───────────────────────────────────────────────────────────

st.markdown('<div class="section-head"><span>03</span> Process</div>', unsafe_allow_html=True)

if st.button("⚡  Process & Generate Output", type="primary", use_container_width=True):

    progress = st.progress(0, text="Reading columns…")

    col_map = {
        "emp_id":       emp_col,
        "emp_name":     name_col,
        "when_change":  when_col,
        "fixed_ctc":    fixed_col,
        "total_pay":    totpay_col,
        "variable_pay": var_col,
        "total_ctc":    total_col,
        "perf_var":     perf_col_r,
        "ret_bonus":    ret_col_r,
        "perf_inc":     perfinc_col_r,
        "status":       status_col,
        "created_date": created_col,
    }

    progress.progress(20, text="Transforming records…")
    result_df = build_employee_records(df, col_map, dayfirst, auto_clean, date_cutoff=filter_date)

    progress.progress(65, text="Building Excel output…")
    xlsx_bytes = to_excel_bytes_output(result_df)

    progress.progress(100, text="Done!")
    progress.empty()

    st.session_state["xlsx_bytes"] = xlsx_bytes
    st.session_state["result_df"]  = result_df
    st.session_state["total_rows"] = len(df)

# ── RESULTS ───────────────────────────────────────────────────────────────────

if "result_df" in st.session_state:
    result_df  = st.session_state["result_df"]
    total_rows = st.session_state["total_rows"]
    n_emp      = len(result_df)

    st.divider()
    st.markdown(f"""
    <div class="metric-row">
      <div class="metric-card"><div class="metric-val">{total_rows:,}</div><div class="metric-label">Input Rows</div></div>
      <div class="metric-card"><div class="metric-val">{n_emp:,}</div><div class="metric-label">Unique Employees</div></div>
      <div class="metric-card"><div class="metric-val">3</div><div class="metric-label">Records Per Employee</div></div>
      <div class="metric-card"><div class="metric-val">29</div><div class="metric-label">Output Columns</div></div>
    </div>
    """, unsafe_allow_html=True)

    st.success(f"✅ Done! {n_emp:,} employees · Latest, 2nd Latest & 3rd Latest records side-by-side")

    st.markdown('<div class="section-head"><span>04</span> Preview (first 10 rows)</div>', unsafe_allow_html=True)
    st.dataframe(result_df.head(10), use_container_width=True, hide_index=True)
    st.caption("ℹ️ Preview shows flat column names. The downloaded Excel has proper merged group headers with colour bands.")

    st.markdown('<div class="section-head"><span>05</span> Download</div>', unsafe_allow_html=True)
    st.download_button(
        "⬇ Download Excel (.xlsx) — Output Sheet Format",
        data=st.session_state["xlsx_bytes"],
        file_name="employee_ctc_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        key="dl_xlsx"
    )
