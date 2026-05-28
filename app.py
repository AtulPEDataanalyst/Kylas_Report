"""
Lead Activity Report Automation Tool — PolicyEra Edition
=========================================================
Production-ready Streamlit app for lead activity processing,
pivot reporting and Excel export.
"""

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import date
import warnings

warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="PolicyEra – Lead Activity Report",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# POLICYERA BRAND COLORS
#   Primary Navy  : #0B1F4B   (deep trust blue)
#   Secondary Navy: #112255
#   Accent Orange : #F47B20   (energy / CTA)
#   Accent Gold   : #FFA940
#   Light BG      : #F4F6FB
#   Card BG       : #FFFFFF
#   Border        : #D8E2F3
#   Text Primary  : #0B1F4B
#   Text Muted    : #6B7A99
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Sora:wght@600;700;800&display=swap');

:root {
    --pe-navy:   #0B1F4B;
    --pe-navy2:  #112255;
    --pe-orange: #F47B20;
    --pe-gold:   #FFA940;
    --pe-light:  #F0F4FA;
    --pe-white:  #FFFFFF;
    --pe-border: #C8D8F0;
    --pe-muted:  #5A6A8A;
}

/* ── Global fonts & background ── */
html, body { font-family: 'Inter', sans-serif !important; }
[class*="stApp"], .stApp,
[data-testid="stAppViewContainer"],
[data-testid="stAppViewContainer"] > .main {
    background-color: var(--pe-light) !important;
}
[data-testid="stAppViewContainer"] > .main .block-container {
    padding-top: 1.5rem;
    max-width: 1440px;
}
/* text defaults */
p, span, label, div { color: #1a2a4a; }

/* ── Sidebar ── */
[data-testid="stSidebar"],
[data-testid="stSidebar"] > div:first-child {
    background-color: var(--pe-navy) !important;
    border-right: 3px solid var(--pe-orange) !important;
}
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] span,
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] div,
[data-testid="stSidebar"] .stMarkdown { color: #dce6f8 !important; }
[data-testid="stSidebar"] hr { border-color: #1e3470 !important; }
/* sidebar multiselect chips */
[data-testid="stSidebar"] [data-baseweb="tag"] {
    background-color: var(--pe-orange) !important;
    color: white !important;
}

/* ── Top header/toolbar ── */
[data-testid="stHeader"] { background: var(--pe-navy) !important; }

/* ── Metric cards ── */
[data-testid="stMetric"] {
    background: var(--pe-white) !important;
    border: 1.5px solid var(--pe-border) !important;
    border-top: 4px solid var(--pe-orange) !important;
    border-radius: 12px !important;
    padding: 1.1rem 1.3rem !important;
    box-shadow: 0 2px 14px rgba(11,31,75,0.08) !important;
}
[data-testid="stMetricLabel"] > div { color: var(--pe-muted) !important; font-size:0.7rem; text-transform:uppercase; letter-spacing:0.1em; font-weight:600; }
[data-testid="stMetricValue"] > div { color: var(--pe-navy) !important; font-family:'Sora',sans-serif; font-size:1.8rem; font-weight:700; }

/* ── Buttons ── */
.stDownloadButton > button,
.stButton > button,
button[kind="primary"] {
    background: var(--pe-orange) !important;
    color: white !important;
    border: none !important;
    border-radius: 8px !important;
    font-family: 'Inter', sans-serif !important;
    font-weight: 600 !important;
    letter-spacing: 0.02em !important;
    padding: 0.55rem 1.5rem !important;
    transition: background 0.2s ease, box-shadow 0.2s ease !important;
}
.stDownloadButton > button:hover,
.stButton > button:hover {
    background: #d06010 !important;
    box-shadow: 0 4px 18px rgba(244,123,32,0.45) !important;
}

/* ── Dataframe container ── */
[data-testid="stDataFrame"] {
    border: 1.5px solid var(--pe-border) !important;
    border-radius: 10px !important;
    overflow: hidden !important;
    box-shadow: 0 2px 10px rgba(11,31,75,0.06) !important;
}
/* dataframe header row */
[data-testid="stDataFrame"] th {
    background-color: var(--pe-navy) !important;
    color: white !important;
    font-weight: 600 !important;
}

/* ── Expander ── */
[data-testid="stExpander"] {
    border: 1.5px solid var(--pe-border) !important;
    border-radius: 10px !important;
    background: var(--pe-white) !important;
    box-shadow: 0 1px 6px rgba(11,31,75,0.05) !important;
}
[data-testid="stExpander"] summary { color: var(--pe-navy) !important; font-weight: 600; }

/* ── Progress bar ── */
[data-testid="stProgressBar"] > div > div,
.stProgress > div > div {
    background: linear-gradient(90deg, var(--pe-navy), var(--pe-orange)) !important;
    border-radius: 4px !important;
}

/* ── File uploader ── */
[data-testid="stFileUploader"] {
    border: 2px dashed var(--pe-orange) !important;
    border-radius: 12px !important;
    background: #fff9f5 !important;
}
[data-testid="stFileUploader"] label { color: var(--pe-navy) !important; }

/* ── Select / multiselect tags ── */
[data-baseweb="tag"] {
    background-color: rgba(11,31,75,0.12) !important;
    color: var(--pe-navy) !important;
    border-radius: 4px !important;
}

/* ── HR ── */
hr { border-color: var(--pe-border) !important; }

/* ── Alerts / info boxes ── */
[data-testid="stAlert"] { border-radius: 8px !important; }

/* ── Custom classes ── */
.pe-hero {
    background: linear-gradient(120deg, var(--pe-navy) 0%, #1a3578 60%, #1e3d8f 100%);
    border-radius: 16px;
    padding: 2rem 2.5rem;
    margin-bottom: 1.5rem;
    border-left: 6px solid var(--pe-orange);
    box-shadow: 0 4px 24px rgba(11,31,75,0.18);
}
.pe-hero-title {
    font-family: 'Sora', sans-serif;
    font-size: 2rem;
    font-weight: 800;
    color: #ffffff;
    line-height: 1.2;
    margin: 0 0 0.4rem 0;
}
.pe-hero-sub {
    color: #a0b8e8;
    font-size: 0.95rem;
    margin: 0;
}
.pe-badge {
    display: inline-block;
    background: rgba(244,123,32,0.15);
    color: var(--pe-orange);
    border: 1px solid rgba(244,123,32,0.4);
    border-radius: 20px;
    padding: 3px 12px;
    font-size: 0.7rem;
    font-weight: 700;
    letter-spacing: 0.08em;
    margin-right: 5px;
}
.section-label {
    font-size: 0.7rem;
    text-transform: uppercase;
    letter-spacing: 0.14em;
    color: var(--pe-muted);
    font-weight: 700;
    margin-bottom: 0.5rem;
}
.pivot-info {
    background: linear-gradient(135deg, #fff8f3 0%, #fff3e8 100%);
    border: 1.5px solid rgba(244,123,32,0.3);
    border-left: 4px solid var(--pe-orange);
    border-radius: 10px;
    padding: 0.9rem 1.2rem;
    font-size: 0.85rem;
    color: var(--pe-navy);
    margin-bottom: 1rem;
}
.success-banner {
    background: linear-gradient(90deg, #e6f9f0 0%, #f0fdf4 100%);
    border: 1.5px solid #6fcf97;
    border-left: 5px solid #27ae60;
    border-radius: 10px;
    padding: 0.8rem 1.2rem;
    color: #1a5e36;
    font-weight: 600;
    font-size: 0.9rem;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────────────────────

# FIX 2: robust exclusion — store as a set, strip + lower-case for comparison
_RAW_EXCLUDED = [
    "Closed Unqualified",
    "Qualified Follow Up",
    "Closed Lost",
    "Old Payment Done",
    "PLVC Completed",
    "Payment Done",
    "Renewal Payment Done",
]
# Deduplicate and preserve for display; comparison done case-insensitively
EXCLUDED_PIPELINE_STAGES_DEFAULT = list(dict.fromkeys(_RAW_EXCLUDED))

DAY_BUCKETS = [
    (0,  3,  "0–3 Days"),
    (4,  7,  "4–7 Days"),
    (8,  15, "8–15 Days"),
    (16, 31, "16–31 Days"),
    (32, 60, "1–2 Months"),
]
DAY_BUCKET_OVERFLOW = "More than 2 Months"
BUCKET_ORDER = ["0–3 Days", "4–7 Days", "8–15 Days", "16–31 Days", "1–2 Months", "More than 2 Months"]

COLUMN_ALIASES = {
    "Pipeline Stage":  ["pipeline stage", "stage", "pipeline_stage", "deal stage"],
    "Owner":           ["owner", "lead owner", "assigned to", "sales owner", "rep"],
    "Task Due On":     ["task due on", "task due", "due date", "task_due_on", "due on"],
    "Latest Activity": ["latest activity on", "latest activity", "last activity",
                        "last_activity", "latest_activity", "activity date"],
}

# ─────────────────────────────────────────────────────────────────────────────
# UTILITY FUNCTIONS
# ─────────────────────────────────────────────────────────────────────────────

def auto_map_columns(df: pd.DataFrame) -> dict:
    col_lower = {c.lower().strip(): c for c in df.columns}
    mapping = {}
    for canonical, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            if alias.lower() in col_lower:
                mapping[canonical] = col_lower[alias.lower()]
                break
    return mapping


def load_file(uploaded_file) -> pd.DataFrame:
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        df = pd.read_csv(uploaded_file, dtype=str)
    elif name.endswith((".xlsx", ".xls")):
        xf = pd.ExcelFile(uploaded_file)
        df = pd.read_excel(xf, sheet_name=xf.sheet_names[0], dtype=str)
    else:
        raise ValueError("Unsupported format. Upload .xlsx or .csv")
    df.replace({"": np.nan, "nan": np.nan, "NaT": np.nan}, inplace=True)
    df.columns = df.columns.str.strip()
    return df


def safe_to_date(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors="coerce", dayfirst=False)


def filter_pipeline(df: pd.DataFrame, stage_col: str, excluded_stages: list) -> pd.DataFrame:
    """
    FIX 2: Case-insensitive exclusion so 'Closed Lost' and 'closed lost' both match.
    """
    excluded_lower = {s.strip().lower() for s in excluded_stages}
    mask = ~df[stage_col].fillna("").str.strip().str.lower().isin(excluded_lower)
    return df[mask].copy().reset_index(drop=True)


def add_future_task(df: pd.DataFrame, due_col: str) -> pd.DataFrame:
    today = pd.Timestamp(date.today())
    dt = safe_to_date(df[due_col])
    future_task = np.where(dt.isna(), "No", np.where(dt >= today, "Yes", "No"))
    idx = df.columns.get_loc(due_col) + 1
    df.insert(idx, "Future Task", future_task)
    return df


def add_days(df: pd.DataFrame, activity_col: str) -> pd.DataFrame:
    today = pd.Timestamp(date.today())
    dt = safe_to_date(df[activity_col])
    days = (today - dt).dt.days.where(dt.notna(), other=np.nan)
    idx = df.columns.get_loc(activity_col) + 1
    df.insert(idx, "Days", days.astype("Int64"))
    return df


def bucket_days(days_val) -> str:
    if pd.isna(days_val):
        return "Unknown"
    d = int(days_val)
    if d < 0:
        return "0–3 Days"
    for lo, hi, label in DAY_BUCKETS:
        if lo <= d <= hi:
            return label
    return DAY_BUCKET_OVERFLOW


def add_group_of_days(df: pd.DataFrame) -> pd.DataFrame:
    group = df["Days"].apply(bucket_days)
    idx = df.columns.get_loc("Days") + 1
    df.insert(idx, "Group of Days", group)
    return df


# ─────────────────────────────────────────────────────────────────────────────
# FIX 1: PIVOT — Owner (grouped) × Pipeline Stage rows, Day Buckets as columns
# ─────────────────────────────────────────────────────────────────────────────
def build_pivot(df: pd.DataFrame, owner_col: str, stage_col: str) -> pd.DataFrame:
    """
    Pivot filtered to Future Task = No.
    Rows : Owner  (group header) + Pipeline Stage (sub-rows)
    Cols : Day bucket columns + Sub-Total per Owner + Grand Total
    """
    sub = df[df["Future Task"] == "No"].copy()

    if sub.empty:
        cols = [owner_col, stage_col] + BUCKET_ORDER + ["Sub-Total"]
        return pd.DataFrame(columns=cols)

    # Count per owner × stage × bucket
    grp = (
        sub.groupby([owner_col, stage_col, "Group of Days"])
        .size()
        .unstack("Group of Days", fill_value=0)
    )

    # Ensure all bucket columns present (ordered)
    for b in BUCKET_ORDER:
        if b not in grp.columns:
            grp[b] = 0
    grp = grp[BUCKET_ORDER]

    # Sub-total per Owner × Stage row
    grp["Sub-Total"] = grp.sum(axis=1)
    grp.reset_index(inplace=True)
    grp.columns.name = None

    # ── Build display dataframe with Owner group rows ──────────────
    # Each Owner gets a summary row (sum across all its stages) + individual stage rows
    bucket_cols = BUCKET_ORDER + ["Sub-Total"]
    result_rows = []

    for owner, owner_df in grp.groupby(owner_col, sort=True):
        # Owner-level summary row
        owner_totals = owner_df[bucket_cols].sum().to_dict()
        summary_row = {owner_col: owner, stage_col: "── TOTAL"}
        summary_row.update(owner_totals)
        summary_row["_row_type"] = "owner_total"
        result_rows.append(summary_row)

        # Stage sub-rows
        for _, row in owner_df.iterrows():
            r = {owner_col: "", stage_col: row[stage_col]}
            r.update({c: row[c] for c in bucket_cols})
            r["_row_type"] = "stage"
            result_rows.append(r)

    # Grand total row
    grand = grp[bucket_cols].sum().to_dict()
    grand_row = {owner_col: "GRAND TOTAL", stage_col: ""}
    grand_row.update(grand)
    grand_row["_row_type"] = "grand_total"
    result_rows.append(grand_row)

    result = pd.DataFrame(result_rows)
    # Reorder columns cleanly
    display_cols = [owner_col, stage_col] + bucket_cols
    result = result[display_cols]
    return result


def build_summary(df: pd.DataFrame, owner_col: str, stage_col: str):
    rows = [
        {"Metric": "Total Leads",        "Value": len(df)},
        {"Metric": "Future Tasks (Yes)",  "Value": int((df["Future Task"] == "Yes").sum())},
        {"Metric": "Pending Tasks (No)",  "Value": int((df["Future Task"] == "No").sum())},
    ]
    owner_counts = df[owner_col].value_counts().reset_index()
    owner_counts.columns = ["Owner", "Lead Count"]
    stage_counts = df[stage_col].value_counts().reset_index()
    stage_counts.columns = ["Pipeline Stage", "Lead Count"]
    return pd.DataFrame(rows), owner_counts, stage_counts


# ─────────────────────────────────────────────────────────────────────────────
# EXCEL EXPORT  — PolicyEra branded colours
# ─────────────────────────────────────────────────────────────────────────────
PE_NAVY   = "#0B1F4B"
PE_NAVY2  = "#112255"
PE_ORANGE = "#F47B20"
PE_GOLD   = "#FFA940"
PE_LIGHT  = "#F4F6FB"
PE_WHITE  = "#FFFFFF"
PE_BORDER = "#D8E2F3"
PE_MUTED  = "#6B7A99"


def to_excel(processed_df, pivot_df, summary_top, owner_counts, stage_counts) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        wb = writer.book

        # ── Shared formats ────────────────────────────────────────────
        def fmt(**kw):
            defaults = {"font_name": "Calibri", "font_size": 10, "valign": "vcenter"}
            defaults.update(kw)
            return wb.add_format(defaults)

        hdr_fmt = fmt(bold=True, bg_color=PE_NAVY, font_color="#FFFFFF",
                      border=1, border_color=PE_NAVY2, align="center", font_size=10)
        cell_fmt  = fmt(border=1, border_color=PE_BORDER)
        alt_fmt   = fmt(border=1, border_color=PE_BORDER, bg_color=PE_LIGHT)
        title_fmt = fmt(bold=True, font_size=14, font_color=PE_NAVY, font_name="Calibri")

        # Pivot-specific row formats
        owner_hdr_fmt = fmt(bold=True, bg_color=PE_NAVY2, font_color="#FFFFFF",
                            border=1, border_color=PE_NAVY, font_size=10)
        stage_fmt     = fmt(border=1, border_color=PE_BORDER, indent=1)
        stage_alt_fmt = fmt(border=1, border_color=PE_BORDER, bg_color=PE_LIGHT, indent=1)
        grand_fmt     = fmt(bold=True, bg_color=PE_ORANGE, font_color="#FFFFFF",
                            border=1, border_color="#c96010", font_size=10)
        num_fmt       = fmt(border=1, border_color=PE_BORDER, align="center", num_format="0")
        num_alt_fmt   = fmt(border=1, border_color=PE_BORDER, bg_color=PE_LIGHT,
                            align="center", num_format="0")
        num_own_fmt   = fmt(bold=True, bg_color=PE_NAVY2, font_color="#FFFFFF",
                            border=1, border_color=PE_NAVY, align="center", num_format="0")
        num_grand_fmt = fmt(bold=True, bg_color=PE_ORANGE, font_color="#FFFFFF",
                            border=1, border_color="#c96010", align="center", num_format="0")
        metric_fmt    = fmt(bold=True, bg_color="#FFF3E8", font_color=PE_NAVY,
                            border=1, border_color="#f4c090", font_size=10)

        # ── Sheet 1: Processed Data ───────────────────────────────────
        ws1 = wb.add_worksheet("Processed Data")
        writer.sheets["Processed Data"] = ws1
        ws1.write(0, 0, "PolicyEra – Lead Activity Processed Data", title_fmt)
        sr = 2
        processed_df.to_excel(writer, sheet_name="Processed Data", index=False, startrow=sr)
        for ci, cn in enumerate(processed_df.columns):
            ws1.write(sr, ci, cn, hdr_fmt)
            w = max(len(str(cn)) + 2,
                    processed_df[cn].astype(str).str.len().max() + 2 if len(processed_df) else 12)
            ws1.set_column(ci, ci, min(w, 38))
        for ri in range(len(processed_df)):
            rfmt = alt_fmt if ri % 2 == 0 else cell_fmt
            for ci in range(len(processed_df.columns)):
                v = processed_df.iloc[ri, ci]
                ws1.write(ri + sr + 1, ci, "" if pd.isna(v) else v, rfmt)

        # ── Sheet 2: Pivot Report (grouped, colour-coded) ─────────────
        ws2 = wb.add_worksheet("Pivot Report")
        writer.sheets["Pivot Report"] = ws2
        ws2.write(0, 0, "PolicyEra – Lead Activity Pivot (Future Task = No)", title_fmt)
        ws2.write(1, 0, f"Owner × Pipeline Stage  |  Activity Buckets  |  Generated: {date.today()}",
                  fmt(italic=True, font_color=PE_MUTED, font_size=9))
        sr2 = 3

        pivot_cols = list(pivot_df.columns)
        # Identify column positions
        owner_ci  = pivot_cols.index(pivot_df.columns[0])
        stage_ci  = pivot_cols.index(pivot_df.columns[1])
        num_start = 2  # bucket columns start at col index 2

        # Header row
        for ci, cn in enumerate(pivot_cols):
            ws2.write(sr2, ci, cn, hdr_fmt)
            ws2.set_column(ci, ci, 18 if ci >= num_start else 22)

        stage_count = 0
        for ri, row in pivot_df.iterrows():
            xr = ri + sr2 + 1
            row_type = "grand" if row[pivot_cols[0]] == "GRAND TOTAL" else \
                       ("owner" if row[pivot_cols[1]] == "── TOTAL" else "stage")

            if row_type == "owner":
                tf, nf = owner_hdr_fmt, num_own_fmt
            elif row_type == "grand":
                tf, nf = grand_fmt, num_grand_fmt
            else:
                stage_count += 1
                tf = stage_fmt if stage_count % 2 == 1 else stage_alt_fmt
                nf = num_fmt   if stage_count % 2 == 1 else num_alt_fmt

            for ci, cn in enumerate(pivot_cols):
                v = row[cn]
                if pd.isna(v):
                    v = ""
                if ci < num_start:
                    ws2.write(xr, ci, v, tf)
                else:
                    ws2.write(xr, ci, v if v != "" else 0, nf)

        # ── Sheet 3: Summary Statistics ────────────────────────────────
        ws3 = wb.add_worksheet("Summary Statistics")
        writer.sheets["Summary Statistics"] = ws3
        ws3.write(0, 0, "PolicyEra – Lead Activity Summary Statistics", title_fmt)
        ws3.set_column(0, 0, 30)
        ws3.set_column(1, 1, 18)

        r = 2
        for _, row in summary_top.iterrows():
            ws3.write(r, 0, row["Metric"], metric_fmt)
            ws3.write(r, 1, row["Value"], cell_fmt)
            r += 1

        r += 1
        ws3.write(r, 0, "Owner-wise Lead Count", title_fmt); r += 1
        ws3.write(r, 0, "Owner", hdr_fmt); ws3.write(r, 1, "Lead Count", hdr_fmt); r += 1
        for i, (_, row) in enumerate(owner_counts.iterrows()):
            rf = alt_fmt if i % 2 == 0 else cell_fmt
            ws3.write(r, 0, row["Owner"], rf)
            ws3.write(r, 1, row["Lead Count"], rf)
            r += 1

        r += 1
        ws3.write(r, 0, "Pipeline Stage Count", title_fmt); r += 1
        ws3.write(r, 0, "Pipeline Stage", hdr_fmt); ws3.write(r, 1, "Lead Count", hdr_fmt); r += 1
        for i, (_, row) in enumerate(stage_counts.iterrows()):
            rf = alt_fmt if i % 2 == 0 else cell_fmt
            ws3.write(r, 0, row["Pipeline Stage"], rf)
            ws3.write(r, 1, row["Lead Count"], rf)
            r += 1

    return output.getvalue()


# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE — initialise before any widget
# ─────────────────────────────────────────────────────────────────────────────
if "filter_owner_options" not in st.session_state:
    st.session_state["filter_owner_options"] = []
if "filter_stage_options" not in st.session_state:
    st.session_state["filter_stage_options"] = []
if "processed" not in st.session_state:
    st.session_state["processed"] = False

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown(
        "<div style='padding:1rem 0 0.5rem 0;'>"
        "<span style='font-family:Sora,sans-serif;font-size:1.3rem;font-weight:800;"
        "color:#F47B20;'>🛡️ PolicyEra</span><br>"
        "<span style='font-size:0.72rem;color:#a0b8e8;letter-spacing:0.08em;'>"
        "LEAD ACTIVITY TOOL</span></div>",
        unsafe_allow_html=True,
    )
    st.markdown("---")

    st.markdown('<p class="section-label">Exclude Pipeline Stages</p>', unsafe_allow_html=True)
    excluded = st.multiselect(
        "Stages to exclude",
        options=sorted(set(EXCLUDED_PIPELINE_STAGES_DEFAULT + (
            st.session_state.get("filter_stage_options", [])))),
        default=EXCLUDED_PIPELINE_STAGES_DEFAULT,
        key="excluded_stages",
    )

    st.markdown("---")
    st.markdown('<p class="section-label">Filter by Owner</p>', unsafe_allow_html=True)
    owner_filter = st.multiselect(
        "Select owners",
        options=st.session_state["filter_owner_options"],
        key="owner_filter",
    )

    st.markdown('<p class="section-label">Filter by Pipeline Stage</p>', unsafe_allow_html=True)
    stage_filter = st.multiselect(
        "Select stages",
        options=st.session_state["filter_stage_options"],
        key="stage_filter",
    )

    st.markdown("---")
    st.markdown('<p class="section-label">Date Range Filter</p>', unsafe_allow_html=True)
    use_date_filter = st.checkbox("Enable date range filter", value=False)
    date_start = date_end = None
    if use_date_filter:
        date_start = st.date_input("Activity from", value=None)
        date_end   = st.date_input("Activity to",   value=None)

    st.markdown("---")
    st.caption("PolicyEra Lead Activity Tool v2.0")

# ─────────────────────────────────────────────────────────────────────────────
# HERO HEADER
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="pe-hero">
  <div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:1rem;">
    <div>
      <p class="pe-hero-title">🛡️ Lead Activity Report</p>
      <p class="pe-hero-sub">Upload &nbsp;→&nbsp; Process &nbsp;→&nbsp; Pivot &nbsp;→&nbsp; Export &nbsp;|&nbsp; PolicyEra CRM Intelligence</p>
    </div>
    <div>
      <span class="pe-badge">XLSX</span>
      <span class="pe-badge">CSV</span>
      <span class="pe-badge">PIVOT</span>
      <span class="pe-badge">CHARTS</span>
      <span class="pe-badge">EXPORT</span>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# FILE UPLOAD
# ─────────────────────────────────────────────────────────────────────────────
st.markdown('<p class="section-label">📂 Upload Lead Activity File</p>', unsafe_allow_html=True)
uploaded_file = st.file_uploader(
    "Drag & drop or click to browse — .xlsx / .csv supported",
    type=["xlsx", "xls", "csv"],
    label_visibility="collapsed",
)

if uploaded_file is not None:
    try:
        raw_df = load_file(uploaded_file)
    except Exception as e:
        st.error(f"❌ Error reading file: {e}")
        st.stop()

    with st.expander("🔍 Raw File Preview", expanded=False):
        st.dataframe(raw_df.head(50), use_container_width=True)
        st.caption(f"Rows: {len(raw_df):,}  •  Columns: {len(raw_df.columns)}")

    # ── Column Mapping ────────────────────────────────────────────────────────
    col_map = auto_map_columns(raw_df)
    missing = [k for k in COLUMN_ALIASES if k not in col_map]
    if missing:
        st.warning(f"⚠️ Auto-detect failed for: **{', '.join(missing)}**. Map them below.")
        col_options = list(raw_df.columns)
        for m in missing:
            chosen = st.selectbox(f"Column for **{m}**",
                                  options=["— Select —"] + col_options, key=f"map_{m}")
            if chosen != "— Select —":
                col_map[m] = chosen

    still_missing = [k for k in COLUMN_ALIASES if k not in col_map]
    if still_missing:
        st.info("Please map all required columns to proceed.")
        st.stop()

    # ── Populate sidebar filter options (safe — never touch widget-bound keys)
    stage_col = col_map["Pipeline Stage"]
    owner_col = col_map["Owner"]
    all_owners = sorted(raw_df[owner_col].dropna().unique().tolist())
    all_stages = sorted(raw_df[stage_col].dropna().unique().tolist())

    if st.session_state["filter_owner_options"] != all_owners:
        st.session_state["filter_owner_options"] = all_owners
       
    if st.session_state["filter_stage_options"] != all_stages:
        st.session_state["filter_stage_options"] = all_stages
        

    # ── Process button ────────────────────────────────────────────────────────
    st.markdown("---")
    process_btn = st.button("🚀  Process Data", use_container_width=False)

    if process_btn or st.session_state.get("processed"):
        st.session_state["processed"] = True

        with st.spinner("Processing…"):
            progress = st.progress(0, text="Loading file…")

            uploaded_file.seek(0)
            df = load_file(uploaded_file)
            progress.progress(15, text="Filtering excluded pipeline stages…")

            # FIX 2: pass the sidebar `excluded` list directly (case-insensitive inside fn)
            df = filter_pipeline(df, col_map["Pipeline Stage"], excluded)
            progress.progress(30, text="Adding Future Task…")

            df = add_future_task(df, col_map["Task Due On"])
            progress.progress(45, text="Calculating Days…")

            df = add_days(df, col_map["Latest Activity"])
            progress.progress(60, text="Bucketing activity age…")

            df = add_group_of_days(df)
            progress.progress(72, text="Applying display filters…")

            if owner_filter:
                df = df[df[owner_col].isin(owner_filter)]
            if stage_filter:
                df = df[df[stage_col].isin(stage_filter)]
            if use_date_filter and date_start and date_end:
                act_col = col_map["Latest Activity"]
                dd = safe_to_date(df[act_col])
                df = df[(dd >= pd.Timestamp(date_start)) & (dd <= pd.Timestamp(date_end))]

            progress.progress(82, text="Building pivot…")
            pivot_df = build_pivot(df, owner_col, stage_col)

            progress.progress(93, text="Compiling statistics…")
            summary_top, owner_counts, stage_counts = build_summary(df, owner_col, stage_col)
            progress.progress(100, text="Done!")

        st.markdown(
            '<div class="success-banner">✅ Data processed successfully!</div>',
            unsafe_allow_html=True,
        )
        st.markdown("<br>", unsafe_allow_html=True)

        # ── KPI Metrics ────────────────────────────────────────────────────
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Total Leads",     f"{len(df):,}")
        m2.metric("Future Tasks ✅",  f"{(df['Future Task']=='Yes').sum():,}")
        m3.metric("Pending Tasks ⏳", f"{(df['Future Task']=='No').sum():,}")
        m4.metric("Unique Owners",    f"{df[owner_col].nunique():,}")
        m5.metric("Pipeline Stages",  f"{df[stage_col].nunique():,}")

        st.markdown("---")

        # ── Processed Data Preview ──────────────────────────────────────────
        with st.expander("📋 Processed Data Preview", expanded=False):
            st.dataframe(df.head(200), use_container_width=True)
            st.caption(f"Showing first 200 of {len(df):,} rows  •  {len(df.columns)} columns")

        # ── Pivot Table ─────────────────────────────────────────────────────
        st.markdown("---")
        st.markdown('<p class="section-label">🔢 Pivot Report — Owner × Stage × Activity Buckets (Future Task = No)</p>',
                    unsafe_allow_html=True)
        st.markdown("""
        <div class="pivot-info">
            <b>How to read:</b>&nbsp; Each <b>owner</b> has a highlighted summary row (── TOTAL)
            followed by their individual pipeline stage rows.
            Columns show lead counts per activity-age bucket.
            Only leads with <b>Future Task = No</b> are counted.
        </div>""", unsafe_allow_html=True)

        # Style the pivot dataframe for display
        def style_pivot(df_p):
            owner_col_name = df_p.columns[0]
            stage_col_name = df_p.columns[1]
            styles = pd.DataFrame("", index=df_p.index, columns=df_p.columns)
            alt_counter = 0
            for pos, (idx, row) in enumerate(df_p.iterrows()):
                if row[owner_col_name] == "GRAND TOTAL":
                    styles.iloc[pos] = (
                        "background-color:#F47B20;color:white;font-weight:700;"
                        "font-size:13px;border-top:2px solid #c96810;"
                    )
                elif str(row[stage_col_name]).strip() == "── TOTAL":
                    styles.iloc[pos] = (
                        "background-color:#0B1F4B;color:white;font-weight:600;"
                        "font-size:12px;"
                    )
                else:
                    alt_counter += 1
                    if alt_counter % 2 == 0:
                        styles.iloc[pos] = "background-color:#EDF1FA;color:#0B1F4B;"
                    else:
                        styles.iloc[pos] = "background-color:#FFFFFF;color:#0B1F4B;"
            return styles

        # reset_index so the integer index is not shown (hide_index not available in older Streamlit)
        pivot_display = pivot_df.reset_index(drop=True)
        st.dataframe(
            pivot_display.style.apply(style_pivot, axis=None),
            use_container_width=True,
        )

        # ── Charts ──────────────────────────────────────────────────────────
        st.markdown("---")
        st.markdown('<p class="section-label">📊 Visual Analysis</p>', unsafe_allow_html=True)

        try:
            import plotly.express as px

            PE_COLORS = ["#0B1F4B", "#1a3578", "#2a4e9e", "#F47B20", "#FFA940", "#f9c87c"]

            c1, c2 = st.columns(2)
            with c1:
                oc = owner_counts.sort_values("Lead Count", ascending=True).tail(15)
                fig1 = px.bar(oc, x="Lead Count", y="Owner", orientation="h",
                              title="Owner-wise Lead Count",
                              color="Lead Count",
                              color_continuous_scale=[[0, "#D8E2F3"], [0.5, "#1a3578"], [1, "#F47B20"]])
                fig1.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="#F4F6FB",
                                   font_color="#0B1F4B", title_font_size=13,
                                   coloraxis_showscale=False,
                                   margin=dict(l=0, r=0, t=35, b=0))
                fig1.update_xaxes(gridcolor="#D8E2F3")
                fig1.update_yaxes(gridcolor="#D8E2F3")
                st.plotly_chart(fig1, use_container_width=True)

            with c2:
                gd = df["Group of Days"].value_counts().reindex(BUCKET_ORDER + ["Unknown"], fill_value=0)
                gd = gd[gd > 0]
                fig2 = px.pie(names=gd.index, values=gd.values,
                              title="Lead Distribution by Activity Age",
                              color_discrete_sequence=PE_COLORS, hole=0.45)
                fig2.update_layout(paper_bgcolor="rgba(0,0,0,0)", font_color="#0B1F4B",
                                   title_font_size=13, margin=dict(l=0, r=0, t=35, b=0))
                st.plotly_chart(fig2, use_container_width=True)

            # Stage × Bucket stacked bar
            stage_grp = (df[df["Future Task"] == "No"]
                         .groupby([stage_col, "Group of Days"])
                         .size().reset_index(name="Count"))
            fig3 = px.bar(stage_grp, x=stage_col, y="Count", color="Group of Days",
                          title="Pipeline Stage × Activity Bucket (Pending Tasks Only)",
                          category_orders={"Group of Days": BUCKET_ORDER},
                          color_discrete_sequence=PE_COLORS)
            fig3.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="#F4F6FB",
                               font_color="#0B1F4B", title_font_size=13,
                               legend_title_text="Activity Bucket",
                               margin=dict(l=0, r=0, t=40, b=0))
            fig3.update_xaxes(gridcolor="#D8E2F3")
            fig3.update_yaxes(gridcolor="#D8E2F3")
            st.plotly_chart(fig3, use_container_width=True)

            # Owner × Bucket heatmap
            heat_data = (df[df["Future Task"] == "No"]
                         .groupby([owner_col, "Group of Days"])
                         .size().reset_index(name="Count"))
            heat_pivot = heat_data.pivot(index=owner_col, columns="Group of Days", values="Count").fillna(0)
            heat_ordered = [b for b in BUCKET_ORDER if b in heat_pivot.columns]
            heat_pivot = heat_pivot[heat_ordered]
            import plotly.graph_objects as go
            fig4 = go.Figure(data=go.Heatmap(
                z=heat_pivot.values,
                x=heat_pivot.columns.tolist(),
                y=heat_pivot.index.tolist(),
                colorscale=[[0, "#F4F6FB"], [0.4, "#1a3578"], [1, "#F47B20"]],
                showscale=True,
            ))
            fig4.update_layout(
                title="Owner × Activity Bucket Heatmap (Pending Tasks)",
                paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="#F4F6FB",
                font_color="#0B1F4B", title_font_size=13,
                margin=dict(l=0, r=0, t=40, b=0),
                xaxis_title="Activity Bucket", yaxis_title="Owner",
            )
            st.plotly_chart(fig4, use_container_width=True)

        except ImportError:
            st.info("Install plotly (`pip install plotly`) for interactive charts.")

        # ── Downloads ────────────────────────────────────────────────────────
        st.markdown("---")
        st.markdown('<p class="section-label">⬇️ Download Report</p>', unsafe_allow_html=True)

        excel_bytes = to_excel(df, pivot_df, summary_top, owner_counts, stage_counts)
        d1, d2, d3 = st.columns([2, 1.5, 1.5])
        with d1:
            st.download_button(
                "⬇️ Download Full Excel Report (3 sheets)",
                data=excel_bytes,
                file_name=f"policyera_lead_report_{date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with d2:
            st.download_button(
                "⬇️ Download Processed CSV",
                data=df.to_csv(index=False).encode("utf-8"),
                file_name=f"policyera_leads_{date.today()}.csv",
                mime="text/csv",
                use_container_width=True,
            )
        with d3:
            # Export pivot without the style metadata column
            st.download_button(
                "⬇️ Download Pivot CSV",
                data=pivot_df.to_csv(index=False).encode("utf-8"),
                file_name=f"policyera_pivot_{date.today()}.csv",
                mime="text/csv",
                use_container_width=True,
            )

else:
    # ── Empty state ────────────────────────────────────────────────────────
    st.markdown("""
    <div style="text-align:center;padding:4rem 2rem;background:#fff;border-radius:16px;
                border:2px dashed #D8E2F3;margin-top:1rem;">
      <div style="font-size:3.5rem;margin-bottom:1rem;">📤</div>
      <div style="font-size:1.1rem;font-weight:700;color:#0B1F4B;">
        Upload your Lead Activity file to get started
      </div>
      <div style="font-size:0.85rem;color:#6B7A99;margin-top:0.5rem;">
        Supports .xlsx and .csv · handles 100,000+ rows
      </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    with st.expander("ℹ️ Required Columns & How It Works", expanded=True):
        ci1, ci2 = st.columns(2)
        with ci1:
            st.markdown("**Required columns (auto-detected):**")
            for col, desc in {
                "Pipeline Stage":  "Filter source — excluded stages are removed",
                "Owner":           "Sales rep / account owner",
                "Task Due On":     "Drives Future Task = Yes / No",
                "Latest Activity": "Drives Days & Activity Bucket columns",
            }.items():
                st.markdown(f"- **`{col}`** — {desc}")
        with ci2:
            st.markdown("**Processing steps:**")
            steps = [
                "Load & auto-detect columns",
                "Remove excluded pipeline stages",
                "Add Future Task (Yes if due ≥ today)",
                "Add Days since latest activity",
                "Add Group of Days bucket",
                "Pivot: Owner → Stage → Buckets",
                "Export 3-sheet branded Excel",
            ]
            for i, s in enumerate(steps, 1):
                st.markdown(f"{i}. {s}")