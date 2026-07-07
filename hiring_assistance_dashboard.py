"""
Hiring Assistance Dashboard
============================
A Streamlit app for CV-screening trackers (Shell / Qlik Sense style sheets).

Run with:  streamlit run hiring_assistance_dashboard.py
"""

import re
from io import BytesIO

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────────────────────────────────────
# Page config & design tokens
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Hiring Assistance Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

NAVY      = "#0F2D5E"
BLUE      = "#1F6FBB"
LIGHT_BLUE= "#EAF2FB"
TEAL      = "#0D8A72"
RED       = "#C0392B"
AMBER     = "#D4870A"
BG        = "#F4F7FB"
SURFACE   = "#FFFFFF"
BORDER    = "#DDE3EE"
TEXT      = "#1A1A2E"
MUTED     = "#6B7A99"

# Chart palette
PALETTE   = [BLUE, TEAL, AMBER, "#7B61FF", "#E87040", "#2EBFA5", "#C0392B"]

st.markdown(
    f"""
    <style>
        /* ── Global background ── */
        .stApp {{ background-color: {BG}; }}

        /* ── Hide default Streamlit chrome ── */
        #footer {{ visibility: hidden; }}

        /* ── Sidebar ── */
        section[data-testid="stSidebar"] {{
            background-color: {NAVY};
        }}
        section[data-testid="stSidebar"] * {{
            color: #E8EEF8 !important;
        }}
        section[data-testid="stSidebar"] .stSlider > div > div > div {{
            background: {BLUE} !important;
        }}
        section[data-testid="stSidebar"] label {{
            font-weight: 600 !important;
            font-size: 0.82rem !important;
            letter-spacing: 0.04em !important;
            text-transform: uppercase !important;
        }}

        /* ── File uploader: hint, button & uploaded file names/sizes ── */
        section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"] span,
        section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"] small,
        section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"] p {{
            color: #000000 !important;
        }}
        section[data-testid="stSidebar"] [data-testid="stFileUploader"] button span {{
            color: #000000 !important;
        }}
        /* Uploaded file name and size in the file list */
        section[data-testid="stSidebar"] [data-testid="stUploadedFile"] span,
        section[data-testid="stSidebar"] [data-testid="stUploadedFile"] p,
        section[data-testid="stSidebar"] [data-testid="stUploadedFile"] small,
        section[data-testid="stSidebar"] [data-testid="stUploadedFileData"] span,
        section[data-testid="stSidebar"] [data-testid="stUploadedFileData"] p {{
            color: #000000 !important;
        }}
        /* The delete (×) button on each uploaded file */
        section[data-testid="stSidebar"] [data-testid="stFileUploaderDeleteBtn"] button span {{
            color: #000000 !important;
        }}

        /* ── Banner ── */
        .dash-banner {{
            background: linear-gradient(135deg, {NAVY} 0%, {BLUE} 100%);
            border-radius: 14px;
            padding: 28px 36px 22px 36px;
            margin-bottom: 24px;
        }}
        .dash-banner h1 {{
            color: #FFFFFF;
            font-size: 1.95rem;
            font-weight: 800;
            margin: 0 0 4px 0;
            letter-spacing: -0.02em;
        }}
        .dash-banner p {{
            color: #B8CFEE;
            font-size: 0.92rem;
            margin: 0;
        }}

        /* ── KPI cards ── */
        .kpi-card {{
            background: {SURFACE};
            border: 1px solid {BORDER};
            border-radius: 12px;
            padding: 16px 18px 12px;
            height: 100%;
        }}
        .kpi-label {{
            font-size: 0.65rem;
            font-weight: 700;
            color: {MUTED};
            text-transform: uppercase;
            letter-spacing: 0.06em;
            margin-bottom: 5px;
        }}
        .kpi-value {{
            font-size: 1.7rem;
            font-weight: 800;
            color: {NAVY};
            line-height: 1.1;
        }}
        .kpi-value-sm {{
            font-size: 1.4rem;
            font-weight: 800;
            color: {NAVY};
            line-height: 1.1;
        }}
        .kpi-sub {{
            font-size: 0.7rem;
            color: {MUTED};
            margin-top: 3px;
        }}
        .kpi-pill {{
            display: inline-block;
            padding: 2px 9px;
            border-radius: 20px;
            font-size: 0.68rem;
            font-weight: 700;
            margin-top: 5px;
        }}
        .pill-green  {{ background:#D4EDDA; color:#1A6630; }}
        .pill-red    {{ background:#FAD7D4; color:#8B1A14; }}
        .pill-amber  {{ background:#FEF0CC; color:#7D4E00; }}
        .pill-blue   {{ background:{LIGHT_BLUE}; color:{NAVY}; }}

        /* ── Section headers ── */
        .sec-head {{
            display: flex;
            align-items: center;
            gap: 8px;
            font-size: 1rem;
            font-weight: 800;
            color: {NAVY};
            margin: 26px 0 12px 0;
            padding-bottom: 8px;
            border-bottom: 2px solid {BORDER};
        }}
        .sec-head span.badge {{
            background: {LIGHT_BLUE};
            color: {BLUE};
            border-radius: 6px;
            padding: 2px 9px;
            font-size: 0.65rem;
            font-weight: 700;
            letter-spacing: 0.05em;
        }}

        /* ── Filters bar ── */
        .filter-bar {{
            background: {SURFACE};
            border: 1px solid {BORDER};
            border-radius: 12px;
            padding: 16px 18px 10px;
            margin-bottom: 14px;
        }}
        .filter-label {{
            font-size: 0.62rem;
            font-weight: 700;
            color: {MUTED};
            text-transform: uppercase;
            letter-spacing: 0.05em;
            margin-bottom: 4px;
        }}

        /* ── Result badge (for table) ── */
        .badge-accept {{ background:#D4EDDA; color:#1A6630;
                         padding:2px 9px; border-radius:12px; font-size:0.78rem; font-weight:700; }}
        .badge-reject {{ background:#FAD7D4; color:#8B1A14;
                         padding:2px 9px; border-radius:12px; font-size:0.78rem; font-weight:700; }}
        .badge-other  {{ background:#E9ECEF; color:#495057;
                         padding:2px 9px; border-radius:12px; font-size:0.78rem; font-weight:700; }}

        /* ── Download button ── */
        div[data-testid="stDownloadButton"] > button {{
            background: {NAVY};
            color: white;
            border: none;
            border-radius: 8px;
            font-weight: 700;
            font-size: 0.9rem;
            padding: 10px 24px;
        }}
        div[data-testid="stDownloadButton"] > button:hover {{
            background: {BLUE};
        }}

        /* ── Dataframe tweaks ── */
        div[data-testid="stDataFrame"] {{ border-radius: 10px; overflow: hidden; }}

        /* ── Expander ── */
        div[data-testid="stExpander"] {{
            border: 1px solid {BORDER};
            border-radius: 10px;
        }}
    </style>
    """,
    unsafe_allow_html=True,
)

# ─────────────────────────────────────────────────────────────────────────────
# Column detection & data cleaning
# ─────────────────────────────────────────────────────────────────────────────
COLUMN_KEYWORDS = {
    "demand_id": ["pmp", "demand id", "requisition", "req id", "position id"],
    "name":      ["candidate name", "name"],
    "pan":       ["pan number", "pan"],
    "vendor":    ["vendor"],
    "account":   ["account name", "account", "client"],
    "jrs":       ["jrs"],
    "status":    ["screen select reason", "screening status", "status", "screen result", "screen"],
    "remarks":   ["remarks or comments", "remarks", "comments", "reason"],
    "week":      ["screened week", "week"],
}


def normalize(s):
    return re.sub(r"[^a-z0-9]", "", str(s).lower())


def find_col(columns, keywords):
    norm_map = {normalize(c): c for c in columns}
    for kw in sorted(keywords, key=len, reverse=True):
        nkw = normalize(kw)
        for ncol, orig in norm_map.items():
            if nkw == ncol:
                return orig
    for kw in sorted(keywords, key=len, reverse=True):
        nkw = normalize(kw)
        for ncol, orig in norm_map.items():
            if nkw in ncol:
                return orig
    return None


def classify_status(value):
    v = str(value).strip().lower()
    if any(k in v for k in ["select", "accept", "shortlist"]):
        return "Accepted"
    if any(k in v for k in ["reject", "decline"]):
        return "Rejected"
    return "Other"


def _find_header_row(xl, sheet, max_scan=6):
    """Detect the real header row (skips title/blank rows in formatted exports)."""
    all_kws = [kw for kws in COLUMN_KEYWORDS.values() for kw in kws]
    probe = xl.parse(sheet, header=None, nrows=max_scan)
    for idx, row in probe.iterrows():
        row_vals = [normalize(str(v)) for v in row if pd.notna(v)]
        for kw in all_kws:
            nkw = normalize(kw)
            if any(nkw == rv or nkw in rv for rv in row_vals):
                return int(idx)
    return 0


@st.cache_data(show_spinner=False)
def load_and_clean(file_bytes, file_name):
    xl = pd.ExcelFile(BytesIO(file_bytes))
    frames = []
    for sheet in xl.sheet_names:
        header_row = _find_header_row(xl, sheet)
        raw = xl.parse(sheet, header=header_row)
        if raw.empty or raw.shape[1] < 3:
            continue
        col_map = {field: find_col(raw.columns, kws) for field, kws in COLUMN_KEYWORDS.items()}
        if not col_map["vendor"] or not col_map["status"]:
            continue
        df = pd.DataFrame()
        for field, src_col in col_map.items():
            df[field] = raw[src_col] if src_col else np.nan
        df = df.dropna(subset=["vendor"], how="any")
        df = df[df["vendor"].astype(str).str.strip() != ""]
        if df.empty:
            continue
        df["source_file"] = file_name
        frames.append(df)
    if not frames:
        return pd.DataFrame(columns=list(COLUMN_KEYWORDS) + ["source_file"])
    out = pd.concat(frames, ignore_index=True)
    out["jrs"]       = out["jrs"].fillna("Unspecified JRS").astype(str).str.strip()
    out["vendor"]    = out["vendor"].astype(str).str.strip()
    out["demand_id"] = out["demand_id"].fillna("Unspecified").astype(str).str.strip()
    out["result"]    = out["status"].apply(classify_status)
    return out


# ─────────────────────────────────────────────────────────────────────────────
# Excel export
# ─────────────────────────────────────────────────────────────────────────────
def build_excel_report(all_data, vendor_stats_display, top_vendors_display, best_vendor_per_jrs, weekly, has_week):
    WHITE_FONT  = Font(color="FFFFFF", bold=True, name="Arial", size=11)
    HEADER_FILL = PatternFill("solid", fgColor="0F2D5E")
    TITLE_FONT  = Font(bold=True, name="Arial", size=14, color="0F2D5E")
    BODY_FONT   = Font(name="Arial", size=10)
    THIN        = Side(style="thin", color="B7B7B7")
    BORDER_STYLE= Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
    CENTER      = Alignment(horizontal="center", vertical="center")
    GREEN       = "C6EFCE"

    def write_df(ws, df, start_row=1, title=None, highlight_col=None):
        r = start_row
        if title:
            ws.cell(row=r, column=1, value=title).font = TITLE_FONT
            r += 2
        header_row = r
        for j, col in enumerate(df.columns, start=1):
            c = ws.cell(row=header_row, column=j, value=col)
            c.font, c.fill, c.alignment, c.border = WHITE_FONT, HEADER_FILL, CENTER, BORDER_STYLE
        for i, row in enumerate(df.itertuples(index=False), start=1):
            for j, val in enumerate(row, start=1):
                c = ws.cell(row=header_row + i, column=j, value=val)
                c.font, c.border = BODY_FONT, BORDER_STYLE
                c.alignment = CENTER if j > 1 else Alignment(horizontal="left")
                if highlight_col and df.columns[j - 1] == highlight_col:
                    c.fill = PatternFill("solid", fgColor=GREEN)
        for j, col in enumerate(df.columns, start=1):
            max_len = max((len(str(v)) for v in df[col]), default=0) if len(df) else 0
            width = max(len(str(col)), max_len) + 4
            ws.column_dimensions[get_column_letter(j)].width = min(max(width, 12), 55)
        ws.freeze_panes = ws.cell(row=header_row + 1, column=1)
        return header_row + len(df) + 1

    total_demands    = all_data["demand_id"].nunique()
    total_candidates = len(all_data)
    total_accepted   = (all_data["result"] == "Accepted").sum()
    total_rejected   = (all_data["result"] == "Rejected").sum()
    total_other      = (all_data["result"] == "Other").sum()
    rate             = total_accepted / total_candidates if total_candidates else 0
    summary_df = pd.DataFrame(
        [
            ("Total Demands (unique PMP/Demand IDs)", total_demands),
            ("Total Unique JRS Roles",                all_data["jrs"].nunique()),
            ("Total Vendors Engaged",                 all_data["vendor"].nunique()),
            ("Total Candidates Screened",             total_candidates),
            ("Total Accepted",                        total_accepted),
            ("Total Rejected",                        total_rejected),
            ("Total Other/Pending",                   total_other),
            ("Overall Acceptance Rate",               f"{total_accepted}/{total_candidates} ({rate:.1%})"),
        ],
        columns=["Metric", "Value"],
    )

    wb  = Workbook()
    ws  = wb.active
    ws.title = "Summary"
    ws.sheet_view.showGridLines = False
    write_df(ws, summary_df, title="CV Screening — Overall Summary")

    ws2 = wb.create_sheet("Vendor Performance")
    ws2.sheet_view.showGridLines = False
    last_row = write_df(ws2, vendor_stats_display, title="Vendor Performance (All Vendors)")
    if len(vendor_stats_display):
        chart = BarChart()
        chart.title = "Acceptance Rate % by Vendor"
        chart.y_axis.title, chart.x_axis.title = "Acceptance Rate %", "Vendor"
        hr = 3
        data_ref = Reference(ws2, min_col=7, min_row=hr, max_row=hr + len(vendor_stats_display))
        cats_ref = Reference(ws2, min_col=1, min_row=hr + 1, max_row=hr + len(vendor_stats_display))
        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cats_ref)
        chart.height, chart.width = 9, 18
        ws2.add_chart(chart, f"A{last_row + 2}")

    ws3 = wb.create_sheet("Top Vendors")
    ws3.sheet_view.showGridLines = False
    write_df(ws3, top_vendors_display, title="Top Vendors by Acceptance Rate", highlight_col="Rank")

    ws4 = wb.create_sheet("Best Vendor per JRS")
    ws4.sheet_view.showGridLines = False
    write_df(ws4, best_vendor_per_jrs, title="Best-Performing Vendor for Each JRS", highlight_col="Best Vendor")

    if has_week:
        ws5 = wb.create_sheet("Weekly Breakdown")
        ws5.sheet_view.showGridLines = False
        write_df(ws5, weekly, title="Weekly Screening Breakdown")

    ws6 = wb.create_sheet("Raw Data")
    ws6.sheet_view.showGridLines = False
    raw_display = all_data.rename(columns={
        "demand_id": "Demand ID (PMP)", "name": "Candidate Name", "pan": "PAN Number",
        "vendor": "Vendor", "account": "Account", "jrs": "JRS",
        "status": "Screening Status", "result": "Result",
        "remarks": "Remarks", "week": "Screened Week", "source_file": "Source File",
    })
    write_df(ws6, raw_display, title="Combined Raw Data")

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ─────────────────────────────────────────────────────────────────────────────
# Helper: KPI card HTML
# ─────────────────────────────────────────────────────────────────────────────
def kpi_card(label, value, sub=None, pill=None, pill_cls="pill-blue"):
    pill_html = f'<div class="kpi-pill {pill_cls}">{pill}</div>' if pill else ""
    sub_html  = f'<div class="kpi-sub">{sub}</div>' if sub else ""
    return (
        f'<div class="kpi-card">'
        f'  <div class="kpi-label">{label}</div>'
        f'  <div class="kpi-value">{value}</div>'
        f'  {sub_html}{pill_html}'
        f'</div>'
    )


def sec_head(icon, title, badge=None):
    badge_html = f'<span class="badge">{badge}</span>' if badge else ""
    st.markdown(
        f'<div class="sec-head">{icon}&nbsp;{title} {badge_html}</div>',
        unsafe_allow_html=True,
    )


# ─────────────────────────────────────────────────────────────────────────────
# Sidebar
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    # ── Logo / title ──
    st.markdown(
        """
        <div style="padding:10px 0 20px 0; border-bottom:1px solid #2A4A7F; margin-bottom:20px;">
            <div style="font-size:1.15rem; font-weight:800; color:#FFFFFF; letter-spacing:-0.01em;">
                📊 Hiring Dashboard
            </div>
            <div style="font-size:0.75rem; color:#8AAAD4; margin-top:2px;">CV Screening Analytics</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ── Data Source ──
    st.markdown(
        "<div style='font-size:0.65rem; font-weight:700; color:#8AAAD4; "
        "text-transform:uppercase; letter-spacing:0.07em; margin-bottom:8px;'>"
        "Data Source</div>",
        unsafe_allow_html=True,
    )
    uploaded_files = st.file_uploader(
        "Upload tracker file(s)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        help="Upload one or more CV screening tracker Excel files.",
        label_visibility="collapsed",
    )
    if uploaded_files:
        names = ", ".join(f.name for f in uploaded_files)
        count_label = f"{len(uploaded_files)} file{'s' if len(uploaded_files) > 1 else ''} loaded"
        st.markdown(
            f"""
            <div style="background:#1A3D6E; border:1px dashed #3A6AAE; border-radius:8px;
                        padding:10px 12px; color:#B8CFEE; font-size:0.72rem; text-align:center;
                        margin-top:4px; margin-bottom:6px; word-break:break-all;">
                📂 {names}<br>
                <span style="color:#5A7AAA; font-size:0.65rem;">{count_label}</span>
            </div>
            """,
            unsafe_allow_html=True,
        )

    # ── Ranking Settings ──
    st.markdown(
        "<div style='font-size:0.65rem; font-weight:700; color:#8AAAD4; "
        "text-transform:uppercase; letter-spacing:0.07em; margin:18px 0 8px;'>"
        "Ranking Settings</div>",
        unsafe_allow_html=True,
    )
    min_submissions = st.slider("Min. submissions to rank vendor", 1, 10, 2)
    top_n           = st.slider("Top N vendors to highlight",      3, 10, 5)

    st.markdown(
        "<div style='margin-top:32px; padding-top:16px; border-top:1px solid #2A4A7F; "
        "font-size:0.65rem; color:#5A7AAA; text-align:center;'>IBM Hiring Analytics · v2.0</div>",
        unsafe_allow_html=True,
    )

# ─────────────────────────────────────────────────────────────────────────────
# Landing state — no file uploaded
# ─────────────────────────────────────────────────────────────────────────────
if not uploaded_files:
    st.markdown(
        f"""
        <div class="dash-banner">
            <h1>Hiring Assistance Dashboard</h1>
            <p>CV Screening · Vendor Performance · Demand Analytics</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.info(
        "👈 **Get started** — upload one or more CV screening tracker Excel files "
        "using the sidebar on the left.",
        icon="📂",
    )
    st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# Load data
# ─────────────────────────────────────────────────────────────────────────────
with st.spinner("Loading and processing data…"):
    all_data = pd.concat(
        [load_and_clean(f.getvalue(), f.name) for f in uploaded_files],
        ignore_index=True,
    )

if all_data.empty:
    st.error(
        "**No usable candidate rows found.** "
        "Check that your file has Vendor and Status columns.",
        icon="⚠️",
    )
    st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# Banner
# ─────────────────────────────────────────────────────────────────────────────
file_label = uploaded_files[0].name if len(uploaded_files) == 1 else f"{len(uploaded_files)} files loaded"
st.markdown(
    f"""
    <div class="dash-banner">
        <h1>Hiring Assistance Dashboard</h1>
        <p>CV Screening · Vendor Performance · Demand Analytics &nbsp;·&nbsp;
           <span style="color:#7FAADD;">{file_label}</span></p>
    </div>
    """,
    unsafe_allow_html=True,
)

# ─────────────────────────────────────────────────────────────────────────────
# KPI strip
# ─────────────────────────────────────────────────────────────────────────────
total_demands    = all_data["demand_id"].nunique()
total_candidates = len(all_data)
total_accepted   = int((all_data["result"] == "Accepted").sum())
total_rejected   = int((all_data["result"] == "Rejected").sum())
total_other      = int((all_data["result"] == "Other").sum())
overall_rate     = total_accepted / total_candidates if total_candidates else 0
n_vendors        = all_data["vendor"].nunique()
n_jrs            = all_data["jrs"].nunique()

k1, k2, k3, k4, k5 = st.columns(5)
with k1:
    st.markdown(kpi_card("Open Demands", total_demands,
                         pill=f"{n_jrs} JRS roles", pill_cls="pill-blue"), unsafe_allow_html=True)
with k2:
    st.markdown(kpi_card("Candidates Screened", f"{total_candidates:,}",
                         sub=f"Across {total_demands} demand IDs"), unsafe_allow_html=True)
with k3:
    st.markdown(kpi_card("Accepted", f"{total_accepted:,}",
                         pill=f"{overall_rate:.1%} acceptance rate", pill_cls="pill-green"), unsafe_allow_html=True)
with k4:
    st.markdown(kpi_card("Rejected", f"{total_rejected:,}",
                         pill=f"{total_rejected/total_candidates:.1%} of total" if total_candidates else "—",
                         pill_cls="pill-red"), unsafe_allow_html=True)
with k5:
    st.markdown(kpi_card("Vendors Engaged", n_vendors,
                         sub=f"Min. {min_submissions} CVs threshold"), unsafe_allow_html=True)

st.markdown("<div style='margin-top:8px;'></div>", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# Section 1 — Demand Drill-down
# ─────────────────────────────────────────────────────────────────────────────
sec_head("🔍", "Demand Drill-down", badge="Cascading Filters")

with st.container():
    st.markdown('<div class="filter-bar">', unsafe_allow_html=True)
    f1, f2, f3, f4 = st.columns(4)
    with f1:
        st.markdown('<div class="filter-label">Demand ID</div>', unsafe_allow_html=True)
        demand_options  = ["All"] + sorted(all_data["demand_id"].unique().tolist())
        selected_demand = st.selectbox("Demand ID", demand_options, label_visibility="collapsed")
    demand_df = all_data if selected_demand == "All" else all_data[all_data["demand_id"] == selected_demand]

    with f2:
        st.markdown('<div class="filter-label">JRS Role</div>', unsafe_allow_html=True)
        jrs_options  = ["All"] + sorted(demand_df["jrs"].unique().tolist())
        selected_jrs = st.selectbox("JRS Role", jrs_options, label_visibility="collapsed")
    jrs_df = demand_df if selected_jrs == "All" else demand_df[demand_df["jrs"] == selected_jrs]

    with f3:
        st.markdown('<div class="filter-label">Vendor</div>', unsafe_allow_html=True)
        vendor_options  = ["All"] + sorted(jrs_df["vendor"].unique().tolist())
        selected_vendor = st.selectbox("Vendor", vendor_options, label_visibility="collapsed")
    vendor_df = jrs_df if selected_vendor == "All" else jrs_df[jrs_df["vendor"] == selected_vendor]

    with f4:
        st.markdown('<div class="filter-label">Screening Result</div>', unsafe_allow_html=True)
        selected_result = st.selectbox(
            "Screening Result",
            ["All", "Accepted", "Rejected", "Other"],
            label_visibility="collapsed",
        )
    filtered_df = vendor_df if selected_result == "All" else vendor_df[vendor_df["result"] == selected_result]
    st.markdown('</div>', unsafe_allow_html=True)

# Drill-down KPI row — uses smaller value font (kpi-value-sm)
n_acc = int((vendor_df["result"] == "Accepted").sum())
n_rej = int((vendor_df["result"] == "Rejected").sum())
n_oth = int((vendor_df["result"] == "Other").sum())
acc_r = n_acc / len(vendor_df) if len(vendor_df) else 0

def kpi_card_sm(label, value, sub=None, pill=None, pill_cls="pill-blue"):
    pill_html = f'<div class="kpi-pill {pill_cls}">{pill}</div>' if pill else ""
    sub_html  = f'<div class="kpi-sub">{sub}</div>' if sub else ""
    return (
        f'<div class="kpi-card">'
        f'  <div class="kpi-label">{label}</div>'
        f'  <div class="kpi-value-sm">{value}</div>'
        f'  {sub_html}{pill_html}'
        f'</div>'
    )

dm1, dm2, dm3, dm4, dm5 = st.columns(5)
with dm1:
    st.markdown(kpi_card_sm("Candidates in Selection", f"{len(demand_df):,}"), unsafe_allow_html=True)
with dm2:
    st.markdown(kpi_card_sm("JRS Roles in Selection", jrs_df["jrs"].nunique()), unsafe_allow_html=True)
with dm3:
    st.markdown(kpi_card_sm("Accepted (filtered)", n_acc,
                            pill=f"{acc_r:.0%}", pill_cls="pill-green"), unsafe_allow_html=True)
with dm4:
    st.markdown(kpi_card_sm("Rejected (filtered)", n_rej,
                            pill_cls="pill-red",
                            pill=f"{n_rej/len(vendor_df):.0%}" if len(vendor_df) else "—"), unsafe_allow_html=True)
with dm5:
    st.markdown(kpi_card_sm("Pending / Other", n_oth,
                            pill_cls="pill-amber",
                            pill=f"{n_oth/len(vendor_df):.0%}" if len(vendor_df) else "—"), unsafe_allow_html=True)

st.markdown("<div style='margin-top:12px;'></div>", unsafe_allow_html=True)

# Candidate table with styled result column
display_cols = {
    "demand_id": "Demand ID", "name": "Candidate Name", "jrs": "JRS",
    "vendor": "Vendor", "status": "Screening Status", "result": "Result", "remarks": "Remarks",
}
available_cols = [c for c in display_cols if c in filtered_df.columns]
table_df = filtered_df[available_cols].rename(columns=display_cols).reset_index(drop=True)

st.dataframe(
    table_df,
    use_container_width=True,
    hide_index=True,
    column_config={
        "Result": st.column_config.TextColumn("Result"),
        "Demand ID": st.column_config.TextColumn("Demand ID"),
        "Candidate Name": st.column_config.TextColumn("Candidate Name", width="medium"),
    },
)
st.caption(f"{len(filtered_df):,} candidate record(s) shown · filtered from {total_candidates:,} total")

# ─────────────────────────────────────────────────────────────────────────────
# Section 2 — Vendor Performance
# ─────────────────────────────────────────────────────────────────────────────
sec_head("🏢", "Vendor Performance", badge=f"All {n_vendors} vendors")

vendor_stats = (
    all_data.groupby("vendor")
    .agg(
        Total_Submitted=("result", "count"),
        Accepted=("result",  lambda s: (s == "Accepted").sum()),
        Rejected=("result",  lambda s: (s == "Rejected").sum()),
        Other=("result",     lambda s: (s == "Other").sum()),
    )
    .reset_index()
    .rename(columns={"vendor": "Vendor"})
)
vendor_stats["Acceptance Rate %"] = (
    vendor_stats["Accepted"] / vendor_stats["Total_Submitted"] * 100
).round(1)
vendor_stats["Acceptance Rate"] = vendor_stats.apply(
    lambda r: f"{r['Accepted']}/{r['Total_Submitted']}", axis=1
)
vendor_stats = vendor_stats.sort_values("Acceptance Rate %", ascending=False).reset_index(drop=True)
vendor_stats["Meets Threshold"] = vendor_stats["Total_Submitted"] >= min_submissions

top_vendors = vendor_stats.sort_values(
    ["Meets Threshold", "Acceptance Rate %", "Total_Submitted"], ascending=[False, False, False]
).head(top_n).reset_index(drop=True)
top_vendors.insert(0, "Rank", range(1, len(top_vendors) + 1))

vendor_stats_display = vendor_stats[
    ["Vendor", "Total_Submitted", "Accepted", "Rejected", "Other", "Acceptance Rate", "Acceptance Rate %"]
]
top_vendors_display = top_vendors[
    ["Rank", "Vendor", "Total_Submitted", "Accepted", "Acceptance Rate", "Acceptance Rate %"]
]

# Row 1 — full-width bar chart
fig_bar = px.bar(
    vendor_stats_display.sort_values("Acceptance Rate %"),
    x="Acceptance Rate %", y="Vendor",
    orientation="h",
    text="Acceptance Rate",
    color="Acceptance Rate %",
    color_continuous_scale=["#F4CCCC", "#FFE699", "#C6EFCE"],
    labels={"Acceptance Rate %": "Acceptance Rate (%)"},
)
fig_bar.update_traces(textposition="outside", textfont_size=11)
fig_bar.update_layout(
    height=max(320, len(vendor_stats_display) * 38),
    margin=dict(t=10, b=10, l=10, r=60),
    coloraxis_showscale=False,
    plot_bgcolor="white",
    paper_bgcolor="white",
    yaxis=dict(tickfont=dict(size=12)),
    xaxis=dict(gridcolor="#EEEEEE", zeroline=False),
)
st.plotly_chart(fig_bar, use_container_width=True)

# Row 2 — donut (left) + top vendors table (right)
vc2, vc3 = st.columns([1, 2])

with vc2:
    donut_df = pd.DataFrame({
        "Result": ["Accepted", "Rejected", "Other"],
        "Count":  [total_accepted, total_rejected, total_other],
    })
    fig_donut = go.Figure(go.Pie(
        labels=donut_df["Result"],
        values=donut_df["Count"],
        hole=0.58,
        marker_colors=[TEAL, RED, AMBER],
        textinfo="percent",
        textfont_size=12,
        hovertemplate="%{label}: %{value} (%{percent})<extra></extra>",
    ))
    fig_donut.update_layout(
        title=dict(text="Overall Outcome Mix", font=dict(size=13, color=NAVY), x=0.5),
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5, font=dict(size=11)),
        margin=dict(t=40, b=20, l=10, r=10),
        height=320,
        paper_bgcolor="white",
    )
    fig_donut.add_annotation(
        text=f"<b>{overall_rate:.0%}</b><br><span style='font-size:10px'>Accept</span>",
        x=0.5, y=0.5, showarrow=False,
        font=dict(size=16, color=NAVY),
        align="center",
    )
    st.plotly_chart(fig_donut, use_container_width=True)

with vc3:
    st.markdown(f"<div style='font-size:0.8rem; font-weight:700; color:{NAVY}; margin-bottom:8px;'>🏆 Top {top_n} Vendors</div>", unsafe_allow_html=True)
    st.dataframe(top_vendors_display, use_container_width=True, hide_index=True, height=320)

with st.expander("View full vendor performance table"):
    st.dataframe(vendor_stats_display, use_container_width=True, hide_index=True)

# ─────────────────────────────────────────────────────────────────────────────
# Section 3 — Best Vendor per JRS
# ─────────────────────────────────────────────────────────────────────────────
sec_head("🎯", "Best Vendor per JRS Role", badge=f"{all_data['jrs'].nunique()} roles")

jrs_vendor = (
    all_data.groupby(["jrs", "vendor"])
    .agg(
        Total_Submitted=("result", "count"),
        Accepted=("result", lambda s: (s == "Accepted").sum()),
    )
    .reset_index()
)
jrs_vendor["Acceptance Rate %"] = (
    jrs_vendor["Accepted"] / jrs_vendor["Total_Submitted"] * 100
).round(1)
eligible_jv = jrs_vendor[jrs_vendor["Total_Submitted"] >= min_submissions]

best_rows = []
for jrs, grp in jrs_vendor.groupby("jrs"):
    pool = eligible_jv[eligible_jv["jrs"] == jrs]
    if pool.empty:
        pool = grp
    best = pool.sort_values(
        ["Acceptance Rate %", "Total_Submitted", "Accepted"], ascending=[False, False, False]
    ).iloc[0]
    best_rows.append(best)

best_vendor_per_jrs = (
    pd.DataFrame(best_rows)
    .rename(columns={"jrs": "JRS", "vendor": "Best Vendor"})
    .reset_index(drop=True)
)
best_vendor_per_jrs["Acceptance Rate"] = best_vendor_per_jrs.apply(
    lambda r: f"{int(r['Accepted'])}/{int(r['Total_Submitted'])}", axis=1
)
best_vendor_per_jrs = best_vendor_per_jrs[
    ["JRS", "Best Vendor", "Total_Submitted", "Accepted", "Acceptance Rate", "Acceptance Rate %"]
]

bv1, bv2 = st.columns([3, 2])
with bv1:
    st.dataframe(best_vendor_per_jrs, use_container_width=True, hide_index=True)
with bv2:
    fig_jrs = px.bar(
        best_vendor_per_jrs.sort_values("Acceptance Rate %"),
        x="Acceptance Rate %", y="JRS",
        orientation="h",
        color="Best Vendor",
        text="Best Vendor",
        color_discrete_sequence=PALETTE,
    )
    fig_jrs.update_traces(textposition="inside", textfont_size=10, insidetextanchor="middle")
    fig_jrs.update_layout(
        height=max(260, len(best_vendor_per_jrs) * 42),
        margin=dict(t=10, b=10, l=10, r=10),
        showlegend=False,
        plot_bgcolor="white",
        paper_bgcolor="white",
        xaxis=dict(gridcolor="#EEEEEE"),
    )
    st.plotly_chart(fig_jrs, use_container_width=True)

# ─────────────────────────────────────────────────────────────────────────────
# Section 4 — Weekly Trend (conditional)
# ─────────────────────────────────────────────────────────────────────────────
has_week = all_data["week"].notna().any()
weekly   = pd.DataFrame(columns=["Screened Week", "Total_Screened", "Accepted", "Rejected", "Acceptance Rate %"])

if has_week:
    sec_head("📅", "Weekly Screening Trend")
    weekly = (
        all_data.dropna(subset=["week"])
        .groupby("week")
        .agg(
            Total_Screened=("result", "count"),
            Accepted=("result",  lambda s: (s == "Accepted").sum()),
            Rejected=("result",  lambda s: (s == "Rejected").sum()),
        )
        .reset_index()
        .rename(columns={"week": "Screened Week"})
    )
    weekly["Acceptance Rate %"] = (weekly["Accepted"] / weekly["Total_Screened"] * 100).round(1)

    fig_week = go.Figure()
    fig_week.add_trace(go.Scatter(
        x=weekly["Screened Week"], y=weekly["Total_Screened"],
        name="Total Screened", mode="lines+markers",
        line=dict(color=BLUE, width=2.5), marker=dict(size=7),
    ))
    fig_week.add_trace(go.Scatter(
        x=weekly["Screened Week"], y=weekly["Accepted"],
        name="Accepted", mode="lines+markers",
        line=dict(color=TEAL, width=2.5, dash="dot"), marker=dict(size=7),
    ))
    fig_week.add_trace(go.Scatter(
        x=weekly["Screened Week"], y=weekly["Rejected"],
        name="Rejected", mode="lines+markers",
        line=dict(color=RED, width=2.5, dash="dot"), marker=dict(size=7),
    ))
    fig_week.update_layout(
        height=340,
        margin=dict(t=10, b=10),
        plot_bgcolor="white",
        paper_bgcolor="white",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        xaxis=dict(gridcolor="#EEEEEE"),
        yaxis=dict(gridcolor="#EEEEEE"),
    )
    st.plotly_chart(fig_week, use_container_width=True)

    with st.expander("View weekly data table"):
        st.dataframe(weekly, use_container_width=True, hide_index=True)

# ─────────────────────────────────────────────────────────────────────────────
# Section 5 — Export
# ─────────────────────────────────────────────────────────────────────────────
sec_head("⬇️", "Export Report")

exp_col, _ = st.columns([2, 5])
with exp_col:
    excel_buf = build_excel_report(
        all_data, vendor_stats_display, top_vendors_display,
        best_vendor_per_jrs, weekly, has_week,
    )
    st.download_button(
        "⬇️  Download Full Excel Report",
        data=excel_buf,
        file_name="CV_Screening_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.caption("Includes Summary, Vendor Performance, Top Vendors, Best Vendor per JRS, and Raw Data sheets.")
