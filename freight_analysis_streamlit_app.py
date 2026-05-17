# ============================================================
# Freight Analysis Dashboard - Streamlit App
# Sheets Used: "Back Up" and "C&FA-Exp"
# Corrected Version:
# 1. Negative values preserved
# 2. Net Sales calculated correctly from Net Bill
# 3. Sales Return / Credit Note treated as negative
# 4. Month order fixed sequentially
# ============================================================

import re
from io import BytesIO
from typing import List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st


# ============================================================
# Page Config
# ============================================================

st.set_page_config(
    page_title="Freight Analysis Dashboard",
    page_icon="🚚",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ============================================================
# Premium CSS
# ============================================================

st.markdown(
    """
    <style>
        .main {
            background: #f7f9fc;
        }

        .block-container {
            padding-top: 1.2rem;
            padding-bottom: 2rem;
        }

        .hero-box {
            background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 52%, #0f766e 100%);
            color: white;
            padding: 24px 28px;
            border-radius: 22px;
            box-shadow: 0 14px 40px rgba(15, 23, 42, 0.18);
            margin-bottom: 18px;
        }

        .hero-title {
            font-size: 34px;
            font-weight: 800;
            margin-bottom: 4px;
            letter-spacing: -0.5px;
        }

        .hero-subtitle {
            font-size: 15px;
            opacity: 0.92;
            margin-top: 3px;
        }

        .metric-card {
            background: white;
            padding: 17px 18px;
            border-radius: 18px;
            border: 1px solid #e5e7eb;
            box-shadow: 0 8px 28px rgba(15, 23, 42, 0.06);
            min-height: 112px;
        }

        .metric-title {
            color: #64748b;
            font-size: 13px;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: .45px;
        }

        .metric-value {
            color: #0f172a;
            font-size: 24px;
            font-weight: 850;
            margin-top: 8px;
            line-height: 1.15;
        }

        .metric-note {
            color: #64748b;
            font-size: 12px;
            margin-top: 7px;
        }

        .section-title {
            font-size: 22px;
            font-weight: 800;
            color: #0f172a;
            margin-top: 12px;
            margin-bottom: 8px;
        }

        .small-caption {
            color: #64748b;
            font-size: 13px;
            margin-bottom: 8px;
        }

        div[data-testid="stDataFrame"] {
            border-radius: 16px;
            overflow: hidden;
            box-shadow: 0 8px 26px rgba(15, 23, 42, 0.05);
        }

        .stTabs [data-baseweb="tab-list"] {
            gap: 10px;
        }

        .stTabs [data-baseweb="tab"] {
            background: white;
            border-radius: 14px 14px 0 0;
            padding: 12px 18px;
            border: 1px solid #e5e7eb;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# Helper Functions
# ============================================================

def clean_col_name(x) -> str:
    """Clean Excel column headers into stable strings."""
    if pd.isna(x):
        return ""

    text = str(x).replace("\n", " ").strip()
    text = re.sub(r"\s+", " ", text)

    return text


def normalize_key(x) -> str:
    """Create a matching key for transporter/CFA names."""
    if pd.isna(x):
        return ""

    text = str(x).upper().replace("&", "AND")
    text = re.sub(r"[^A-Z0-9]+", "", text)

    return text


def to_number(series: pd.Series) -> pd.Series:
    """
    Convert mixed Excel values into numeric values safely.

    Important correction:
    Negative values are preserved.
    Earlier logic replacing '-' with '0' caused wrong Net Sales.
    """

    s = series.astype(str).str.strip()

    s = (
        s.str.replace(",", "", regex=False)
        .str.replace("₹", "", regex=False)
        .str.replace("(", "-", regex=False)
        .str.replace(")", "", regex=False)
        .str.replace("–", "-", regex=False)
        .str.replace("—", "-", regex=False)
    )

    # Only standalone dash or blanks should become zero.
    # Real negative numbers like -25000 will remain negative.
    s = s.replace(["", "-", "nan", "None", "NaT", "NULL", "null"], "0")

    return pd.to_numeric(s, errors="coerce").fillna(0)


def fmt_inr(value) -> str:
    """Indian number format with rupee sign, no decimal."""
    try:
        value = float(value)
    except Exception:
        return "₹0"

    sign = "-" if value < 0 else ""
    value = abs(round(value))

    s = str(int(value))

    if len(s) <= 3:
        return f"{sign}₹{s}"

    last3 = s[-3:]
    rest = s[:-3]

    parts = []

    while len(rest) > 2:
        parts.insert(0, rest[-2:])
        rest = rest[:-2]

    if rest:
        parts.insert(0, rest)

    return f"{sign}₹{','.join(parts)},{last3}"


def fmt_num(value) -> str:
    try:
        return f"{float(value):,.0f}"
    except Exception:
        return "0"


def safe_div(num, den):
    try:
        den = float(den)

        if den == 0:
            return 0

        return float(num) / den

    except Exception:
        return 0


def find_sheet_name(sheet_names: List[str], target: str) -> Optional[str]:
    """Find sheet name case/space-insensitively."""
    target_key = normalize_key(target)

    for s in sheet_names:
        if normalize_key(s) == target_key:
            return s

    return None


def read_excel_file(uploaded_file, sheet_name: str, header=None) -> pd.DataFrame:
    """Read Excel including .xlsb, .xlsx, .xlsm."""
    file_name = uploaded_file.name.lower()

    uploaded_file.seek(0)

    if file_name.endswith(".xlsb"):
        return pd.read_excel(
            uploaded_file,
            sheet_name=sheet_name,
            engine="pyxlsb",
            header=header
        )

    return pd.read_excel(
        uploaded_file,
        sheet_name=sheet_name,
        header=header
    )


def get_excel_sheets(uploaded_file) -> List[str]:
    uploaded_file.seek(0)

    if uploaded_file.name.lower().endswith(".xlsb"):
        xl = pd.ExcelFile(uploaded_file, engine="pyxlsb")
    else:
        xl = pd.ExcelFile(uploaded_file)

    return xl.sheet_names


def fix_month_order(df: pd.DataFrame, month_col: str = "Month") -> Tuple[pd.DataFrame, List[str]]:
    """
    Fix month sorting like:
    Jan-26, Feb-26, Mar-26, Apr-26
    instead of alphabetical order.
    """

    df = df.copy()

    if month_col not in df.columns:
        return df, []

    month_text = df[month_col].astype(str).str.strip()

    df["_Month_Date"] = pd.to_datetime(
        month_text,
        format="%b-%y",
        errors="coerce"
    )

    if df["_Month_Date"].isna().all():
        df["_Month_Date"] = pd.to_datetime(
            month_text,
            errors="coerce",
            dayfirst=True
        )

    if df["_Month_Date"].isna().all():
        month_map = {
            "jan": 1,
            "feb": 2,
            "mar": 3,
            "apr": 4,
            "may": 5,
            "jun": 6,
            "jul": 7,
            "aug": 8,
            "sep": 9,
            "oct": 10,
            "nov": 11,
            "dec": 12,
        }

        temp_month = month_text.str.lower().str[:3].map(month_map)

        df["_Month_Date"] = pd.to_datetime(
            {
                "year": 2026,
                "month": temp_month.fillna(12),
                "day": 1,
            },
            errors="coerce",
        )

    df = df.sort_values("_Month_Date")

    month_order = (
        df.dropna(subset=["_Month_Date"])
        .sort_values("_Month_Date")[month_col]
        .astype(str)
        .drop_duplicates()
        .tolist()
    )

    if not month_order:
        month_order = df[month_col].astype(str).drop_duplicates().tolist()

    return df, month_order


def make_bar(df, x, y, title, color=None, orientation="v", text_auto=True, category_orders=None):
    if df.empty:
        st.info("No data available for this chart based on current filters.")
        return

    fig = px.bar(
        df,
        x=x,
        y=y,
        title=title,
        color=color,
        orientation=orientation,
        text_auto=text_auto,
        category_orders=category_orders,
    )

    fig.update_layout(
        title_font_size=18,
        margin=dict(l=15, r=15, t=55, b=15),
        height=430,
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        legend_title_text="",
    )

    fig.update_traces(
        textposition="outside",
        cliponaxis=False
    )

    st.plotly_chart(fig, use_container_width=True)


def make_line(df, x, y, title, color=None, category_orders=None):
    if df.empty:
        st.info("No data available for this chart based on current filters.")
        return

    fig = px.line(
        df,
        x=x,
        y=y,
        title=title,
        color=color,
        markers=True,
        category_orders=category_orders,
    )

    fig.update_layout(
        title_font_size=18,
        margin=dict(l=15, r=15, t=55, b=15),
        height=420,
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        legend_title_text="",
    )

    st.plotly_chart(fig, use_container_width=True)


def show_metric_card(title: str, value: str, note: str = ""):
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-title">{title}</div>
            <div class="metric-value">{value}</div>
            <div class="metric-note">{note}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def apply_filters(df: pd.DataFrame, filters: dict) -> pd.DataFrame:
    out = df.copy()

    for col, selected in filters.items():
        if col in out.columns and selected:
            out = out[out[col].astype(str).isin(selected)]

    return out


def multi_filter(label: str, df: pd.DataFrame, col: str):
    if col not in df.columns:
        return []

    options = sorted(
        [
            x
            for x in df[col].dropna().astype(str).unique().tolist()
            if x and x.lower() != "nan"
        ]
    )

    return st.sidebar.multiselect(
        label,
        options=options,
        default=[]
    )


def download_excel_button(sheets: dict, file_name: str):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for sheet, data in sheets.items():
            safe_sheet = sheet[:31]
            data.to_excel(writer, index=False, sheet_name=safe_sheet)

            worksheet = writer.sheets[safe_sheet]

            for idx, col in enumerate(data.columns):
                width = min(max(len(str(col)) + 2, 12), 35)
                worksheet.set_column(idx, idx, width)

    st.download_button(
        label="⬇️ Download Filtered Report in Excel",
        data=output.getvalue(),
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ============================================================
# Data Loading and Preparation
# ============================================================

@st.cache_data(show_spinner=False)
def load_and_prepare(file_bytes: bytes, file_name: str) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Load Back Up and C&FA-Exp sheets, then prepare clean tables."""

    bio = BytesIO(file_bytes)
    bio.name = file_name

    sheet_names = get_excel_sheets(bio)

    backup_sheet = find_sheet_name(sheet_names, "Back Up")
    cfa_sheet = find_sheet_name(sheet_names, "C&FA-Exp")

    if not backup_sheet:
        raise ValueError('Sheet "Back Up" not found in uploaded workbook.')

    if not cfa_sheet:
        raise ValueError('Sheet "C&FA-Exp" not found in uploaded workbook.')

    # ------------------------------------------------------------
    # Back Up sheet
    # Header is on Excel row 2, so header=1.
    # ------------------------------------------------------------

    bio = BytesIO(file_bytes)
    bio.name = file_name

    backup = read_excel_file(
        bio,
        backup_sheet,
        header=1
    )

    backup.columns = [clean_col_name(c) for c in backup.columns]
    backup = backup.loc[:, [c for c in backup.columns if c != ""]]
    backup = backup.dropna(how="all")

    # Keep only transaction rows if Invoice No exists.
    if "Invoice No" in backup.columns:
        backup = backup[backup["Invoice No"].notna()].copy()

    # Date conversion for .xlsb serial dates.
    if "Date" in backup.columns:
        if pd.api.types.is_numeric_dtype(backup["Date"]):
            backup["Date"] = pd.to_datetime(
                backup["Date"],
                unit="D",
                origin="1899-12-30",
                errors="coerce"
            )
        else:
            backup["Date"] = pd.to_datetime(
                backup["Date"],
                errors="coerce",
                dayfirst=True
            )

    # Numeric columns
    numeric_cols = [
        "Net Bill",
        "Total Bill",
        "Box",
        "Weight",
        "Loading",
        "Unloading",
        "Freight",
        "Freight %",
        "Received Bill Amount",
    ]

    for col in numeric_cols:
        if col in backup.columns:
            backup[col] = to_number(backup[col])

    # Text columns
    text_cols = [
        "Month",
        "Plant Name",
        "Billing Type",
        "Sales Verticle",
        "Sales Vertical",
        "State",
        "Zone",
        "Transporter",
        "Expense type",
        "Transporter/CFA",
        "Party Code",
        "TXTSold-to party",
    ]

    for col in text_cols:
        if col in backup.columns:
            backup[col] = (
                backup[col]
                .astype(str)
                .replace("nan", np.nan)
                .str.strip()
            )

    # Standardize Sales Vertical column spelling if required.
    if "Sales Vertical" in backup.columns and "Sales Verticle" not in backup.columns:
        backup["Sales Verticle"] = backup["Sales Vertical"]

    # ------------------------------------------------------------
    # Correct Net Sales Calculation
    # ------------------------------------------------------------

    if "Net Bill" not in backup.columns:
        backup["Net Bill"] = 0

    backup["Net Sales"] = backup["Net Bill"]

    if "Billing Type" in backup.columns:
        billing_type_clean = backup["Billing Type"].astype(str).str.upper().str.strip()

        return_mask = (
            billing_type_clean.str.contains("SALE RETURN", na=False)
            | billing_type_clean.str.contains("SALES RETURN", na=False)
            | billing_type_clean.str.contains("RETURN", na=False)
            | billing_type_clean.str.contains("CREDIT NOTE", na=False)
            | billing_type_clean.str.contains("CREDIT", na=False)
            | billing_type_clean.str.contains("CN", na=False)
        )

        # If returns / credit notes are positive in Excel, make them negative.
        # If already negative, keep them negative.
        backup.loc[return_mask, "Net Sales"] = -backup.loc[return_mask, "Net Bill"].abs()

    # ------------------------------------------------------------
    # Calculated fields
    # ------------------------------------------------------------

    backup["Calculated Freight %"] = np.where(
        backup["Net Sales"].astype(float) != 0,
        backup.get("Freight", 0).astype(float) / backup["Net Sales"].astype(float),
        0,
    )

    backup["Received vs Freight Variance"] = (
        backup.get("Received Bill Amount", 0)
        - backup.get("Freight", 0)
    )

    backup["Loading + Unloading"] = (
        backup.get("Loading", 0)
        + backup.get("Unloading", 0)
    )

    backup["Total Logistic Cost"] = (
        backup.get("Freight", 0)
        + backup["Loading + Unloading"]
    )

    if "Transporter" in backup.columns:
        backup["Transporter Match Key"] = backup["Transporter"].apply(normalize_key)
    else:
        backup["Transporter Match Key"] = ""

    if "Transporter/CFA" in backup.columns:
        backup["Transporter CFA Match Key"] = backup["Transporter/CFA"].apply(normalize_key)
    else:
        backup["Transporter CFA Match Key"] = backup["Transporter Match Key"]

    # ------------------------------------------------------------
    # C&FA-Exp sheet
    # ------------------------------------------------------------

    bio = BytesIO(file_bytes)
    bio.name = file_name

    raw = read_excel_file(
        bio,
        cfa_sheet,
        header=None
    )

    raw = raw.dropna(how="all").reset_index(drop=True)

    # Locate row where first two columns are Month and Nature.
    header_idx = None

    for i in range(len(raw)):
        first = clean_col_name(raw.iloc[i, 0]).upper()
        second = clean_col_name(raw.iloc[i, 1]).upper()

        if first == "MONTH" and second == "NATURE":
            header_idx = i
            break

    if header_idx is None:
        raise ValueError('Could not identify Month/Nature header row in "C&FA-Exp" sheet.')

    cfa_name_row = header_idx - 1

    location_headers = [clean_col_name(x) for x in raw.iloc[header_idx].tolist()]
    cfa_names = [clean_col_name(x) for x in raw.iloc[cfa_name_row].tolist()]

    data = raw.iloc[header_idx + 1:].copy().reset_index(drop=True)
    data.columns = location_headers
    data = data.dropna(how="all")

    month_col = data.columns[0]
    nature_col = data.columns[1]

    data = data.rename(
        columns={
            month_col: "Month",
            nature_col: "Nature"
        }
    )

    data["Month"] = (
        data["Month"]
        .astype(str)
        .replace("nan", np.nan)
        .str.strip()
    )

    data["Nature"] = (
        data["Nature"]
        .astype(str)
        .replace("nan", np.nan)
        .str.strip()
    )

    data = data[
        data["Month"].notna()
        & data["Nature"].notna()
    ].copy()

    long_rows = []

    for col_idx, col in enumerate(data.columns):
        if col in ["Month", "Nature"]:
            continue

        if normalize_key(col) in ["GRANDTOTAL", "TOTAL"]:
            continue

        cfa_name = cfa_names[col_idx] if col_idx < len(cfa_names) else col
        location = col

        temp = data[["Month", "Nature", col]].copy()
        temp = temp.rename(columns={col: "Expense Amount"})
        temp["CFA Name"] = cfa_name
        temp["CFA Location"] = location
        temp["Expense Amount"] = to_number(temp["Expense Amount"])

        long_rows.append(temp)

    if long_rows:
        cfa_long = pd.concat(long_rows, ignore_index=True)
    else:
        cfa_long = pd.DataFrame()

    if not cfa_long.empty:
        cfa_long = cfa_long[cfa_long["Expense Amount"] != 0].copy()
        cfa_long["CFA Match Key"] = cfa_long["CFA Name"].apply(normalize_key)
    else:
        cfa_long = pd.DataFrame(
            columns=[
                "Month",
                "Nature",
                "Expense Amount",
                "CFA Name",
                "CFA Location",
                "CFA Match Key",
            ]
        )

    # ------------------------------------------------------------
    # C&FA Reconciliation
    # ------------------------------------------------------------

    backup_cfa_base = backup.copy()

    if "Transporter/CFA" in backup_cfa_base.columns:
        name_col = "Transporter/CFA"
        key_col = "Transporter CFA Match Key"
    elif "Transporter" in backup_cfa_base.columns:
        name_col = "Transporter"
        key_col = "Transporter Match Key"
    else:
        backup_cfa_base["Transporter"] = ""
        name_col = "Transporter"
        key_col = "Transporter Match Key"

    invoice_agg = ("Invoice No", "nunique") if "Invoice No" in backup.columns else ("Freight", "count")

    backup_cfa_sum = (
        backup_cfa_base.groupby([key_col, name_col], dropna=False)
        .agg(
            Backup_Freight=("Freight", "sum"),
            Backup_Received_Bill=("Received Bill Amount", "sum"),
            Backup_Net_Sales=("Net Sales", "sum"),
            Invoice_Count=invoice_agg,
        )
        .reset_index()
        .rename(
            columns={
                key_col: "Match Key",
                name_col: "Back Up Transporter/CFA"
            }
        )
    )

    cfa_sum = (
        cfa_long.groupby(["CFA Match Key", "CFA Name"], dropna=False)
        .agg(
            CFA_Expense=("Expense Amount", "sum")
        )
        .reset_index()
        .rename(
            columns={
                "CFA Match Key": "Match Key"
            }
        )
    )

    recon = backup_cfa_sum.merge(
        cfa_sum,
        on="Match Key",
        how="outer"
    )

    recon["Back Up Transporter/CFA"] = recon["Back Up Transporter/CFA"].fillna("")
    recon["CFA Name"] = recon["CFA Name"].fillna("")

    for c in [
        "Backup_Freight",
        "Backup_Received_Bill",
        "Backup_Net_Sales",
        "Invoice_Count",
        "CFA_Expense"
    ]:
        recon[c] = recon[c].fillna(0)

    recon["CFA Expense vs Received Bill Variance"] = (
        recon["CFA_Expense"]
        - recon["Backup_Received_Bill"]
    )

    recon["CFA Expense vs Freight Variance"] = (
        recon["CFA_Expense"]
        - recon["Backup_Freight"]
    )

    return backup, cfa_long, recon


# ============================================================
# Header
# ============================================================

st.markdown(
    """
    <div class="hero-box">
        <div class="hero-title">🚚 Freight Analysis & C&FA Expense Control Tower</div>
        <div class="hero-subtitle">
            Analyze Net Sales vs Freight, Received Bill, Sales Vertical, State, Zone, Transporter and C&FA nature-wise expenses from Back Up and C&FA-Exp sheets.
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# Upload
# ============================================================

st.sidebar.header("📂 Upload Workbook")

uploaded_file = st.sidebar.file_uploader(
    "Upload Excel file (.xlsb / .xlsx / .xlsm)",
    type=["xlsb", "xlsx", "xlsm", "xls"],
)

if not uploaded_file:
    st.info("Please upload the Freight Provision workbook containing sheets named **Back Up** and **C&FA-Exp**.")
    st.stop()

try:
    backup_df, cfa_df, recon_df = load_and_prepare(
        uploaded_file.getvalue(),
        uploaded_file.name
    )
except Exception as e:
    st.error(f"Unable to process workbook: {e}")
    st.stop()


# ============================================================
# Sidebar Filters
# ============================================================

st.sidebar.header("🔎 Filters")

filters = {
    "Month": multi_filter("Month", backup_df, "Month"),
    "Sales Verticle": multi_filter("Sales Vertical", backup_df, "Sales Verticle"),
    "Zone": multi_filter("Zone", backup_df, "Zone"),
    "State": multi_filter("State", backup_df, "State"),
    "Plant Name": multi_filter("Plant", backup_df, "Plant Name"),
    "Transporter": multi_filter("Transporter", backup_df, "Transporter"),
    "Expense type": multi_filter("Expense Type", backup_df, "Expense type"),
}

filtered = apply_filters(
    backup_df,
    filters
)

st.sidebar.header("🏢 C&FA Filters")

cfa_months = multi_filter("C&FA Month", cfa_df, "Month")
cfa_names = multi_filter("C&FA Name", cfa_df, "CFA Name")
cfa_natures = multi_filter("Expense Nature", cfa_df, "Nature")

cfa_filtered = apply_filters(
    cfa_df,
    {
        "Month": cfa_months,
        "CFA Name": cfa_names,
        "Nature": cfa_natures,
    }
)


# ============================================================
# KPI Calculation
# ============================================================

st.markdown(
    '<div class="section-title">Executive Summary</div>',
    unsafe_allow_html=True
)

net_sales = filtered["Net Sales"].sum() if "Net Sales" in filtered.columns else 0
raw_net_bill = filtered["Net Bill"].sum() if "Net Bill" in filtered.columns else 0
freight = filtered["Freight"].sum() if "Freight" in filtered.columns else 0
received = filtered["Received Bill Amount"].sum() if "Received Bill Amount" in filtered.columns else 0
loading_unloading = filtered["Loading + Unloading"].sum() if "Loading + Unloading" in filtered.columns else 0
total_log_cost = filtered["Total Logistic Cost"].sum() if "Total Logistic Cost" in filtered.columns else 0

freight_percent = safe_div(freight, net_sales) * 100
variance = received - freight

invoice_count = filtered["Invoice No"].nunique() if "Invoice No" in filtered.columns else len(filtered)
transporter_count = filtered["Transporter"].nunique() if "Transporter" in filtered.columns else 0

cfa_total = cfa_filtered["Expense Amount"].sum() if not cfa_filtered.empty else 0


# ============================================================
# KPI Cards
# ============================================================

k1, k2, k3, k4 = st.columns(4)

with k1:
    show_metric_card(
        "Net Sales",
        fmt_inr(net_sales),
        f"Raw Net Bill: {fmt_inr(raw_net_bill)} | Invoices: {fmt_num(invoice_count)}"
    )

with k2:
    show_metric_card(
        "Freight",
        fmt_inr(freight),
        f"Freight % on Net Sales: {freight_percent:.2f}%"
    )

with k3:
    show_metric_card(
        "Received Bill",
        fmt_inr(received),
        f"Variance vs Freight: {fmt_inr(variance)}"
    )

with k4:
    show_metric_card(
        "Transporters",
        fmt_num(transporter_count),
        f"Loading + Unloading: {fmt_inr(loading_unloading)}"
    )

k5, k6, k7, k8 = st.columns(4)

with k5:
    show_metric_card(
        "Total Logistic Cost",
        fmt_inr(total_log_cost),
        "Freight + Loading + Unloading"
    )

with k6:
    show_metric_card(
        "Total Weight",
        fmt_num(filtered["Weight"].sum() if "Weight" in filtered.columns else 0),
        "As per Back Up sheet"
    )

with k7:
    show_metric_card(
        "Total Boxes",
        fmt_num(filtered["Box"].sum() if "Box" in filtered.columns else 0),
        "As per Back Up sheet"
    )

with k8:
    show_metric_card(
        "C&FA Expense",
        fmt_inr(cfa_total),
        "As per C&FA-Exp sheet"
    )


# ============================================================
# Tabs
# ============================================================

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
    [
        "📊 Freight Dashboard",
        "🚛 Transporter Analysis",
        "🌍 Sales / State / Zone",
        "🏢 C&FA Expense Analysis",
        "🔁 C&FA Reconciliation",
        "📄 Data & Download",
    ]
)


# ============================================================
# Tab 1: Freight Dashboard
# ============================================================

with tab1:
    st.markdown(
        '<div class="section-title">Freight Overview</div>',
        unsafe_allow_html=True
    )

    c1, c2 = st.columns(2)

    if "Month" in filtered.columns:
        monthly = (
            filtered.groupby("Month", dropna=False)
            .agg(
                Net_Sales=("Net Sales", "sum"),
                Freight=("Freight", "sum"),
                Received_Bill=("Received Bill Amount", "sum"),
            )
            .reset_index()
        )

        monthly["Freight %"] = np.where(
            monthly["Net_Sales"] != 0,
            monthly["Freight"] / monthly["Net_Sales"] * 100,
            0
        )

        monthly, month_order = fix_month_order(monthly, "Month")

    else:
        monthly = pd.DataFrame()
        month_order = []

    with c1:
        if not monthly.empty:
            plot_df = monthly.melt(
                id_vars="Month",
                value_vars=["Net_Sales", "Freight", "Received_Bill"],
                var_name="Metric",
                value_name="Amount"
            )

            make_bar(
                plot_df,
                "Month",
                "Amount",
                "Month-wise Net Sales vs Freight vs Received Bill",
                color="Metric",
                category_orders={"Month": month_order}
            )

    with c2:
        if not monthly.empty:
            make_line(
                monthly,
                "Month",
                "Freight %",
                "Month-wise Freight % on Net Sales",
                category_orders={"Month": month_order}
            )

    c3, c4 = st.columns(2)

    with c3:
        top_state = (
            filtered.groupby("State", dropna=False)
            .agg(
                Freight=("Freight", "sum")
            )
            .reset_index()
            .sort_values("Freight", ascending=False)
            .head(15)
        ) if "State" in filtered.columns else pd.DataFrame()

        make_bar(
            top_state,
            "State",
            "Freight",
            "Top States by Freight"
        )

    with c4:
        top_zone = (
            filtered.groupby("Zone", dropna=False)
            .agg(
                Freight=("Freight", "sum")
            )
            .reset_index()
            .sort_values("Freight", ascending=False)
        ) if "Zone" in filtered.columns else pd.DataFrame()

        make_bar(
            top_zone,
            "Zone",
            "Freight",
            "Zone-wise Freight"
        )

    st.markdown("#### Month-wise Summary")
    st.dataframe(
        monthly.drop(columns=["_Month_Date"], errors="ignore"),
        use_container_width=True,
        hide_index=True
    )


# ============================================================
# Tab 2: Transporter Analysis
# ============================================================

with tab2:
    st.markdown(
        '<div class="section-title">Transporter Performance Analysis</div>',
        unsafe_allow_html=True
    )

    if "Transporter" in filtered.columns:
        invoice_agg = ("Invoice No", "nunique") if "Invoice No" in filtered.columns else ("Freight", "count")

        transporter = (
            filtered.groupby("Transporter", dropna=False)
            .agg(
                Net_Sales=("Net Sales", "sum"),
                Freight=("Freight", "sum"),
                Received_Bill=("Received Bill Amount", "sum"),
                Loading=("Loading", "sum"),
                Unloading=("Unloading", "sum"),
                Weight=("Weight", "sum"),
                Box=("Box", "sum"),
                Invoice_Count=invoice_agg,
            )
            .reset_index()
            .sort_values("Freight", ascending=False)
        )
    else:
        transporter = pd.DataFrame()

    if not transporter.empty:
        transporter["Freight %"] = np.where(
            transporter["Net_Sales"] != 0,
            transporter["Freight"] / transporter["Net_Sales"] * 100,
            0
        )

        transporter["Received vs Freight Variance"] = (
            transporter["Received_Bill"]
            - transporter["Freight"]
        )

        transporter["Freight per Kg"] = np.where(
            transporter["Weight"] != 0,
            transporter["Freight"] / transporter["Weight"],
            0
        )

        transporter["Freight per Box"] = np.where(
            transporter["Box"] != 0,
            transporter["Freight"] / transporter["Box"],
            0
        )

        c1, c2 = st.columns(2)

        with c1:
            make_bar(
                transporter.head(15),
                "Transporter",
                "Freight",
                "Top 15 Transporters by Freight"
            )

        with c2:
            make_bar(
                transporter.head(15),
                "Transporter",
                "Received_Bill",
                "Top 15 Transporters by Received Bill"
            )

        c3, c4 = st.columns(2)

        with c3:
            make_bar(
                transporter.sort_values(
                    "Received vs Freight Variance",
                    ascending=False
                ).head(15),
                "Transporter",
                "Received vs Freight Variance",
                "Highest Received Bill Variance"
            )

        with c4:
            make_bar(
                transporter.sort_values(
                    "Freight %",
                    ascending=False
                ).head(15),
                "Transporter",
                "Freight %",
                "Highest Freight % Transporters"
            )

        st.dataframe(
            transporter,
            use_container_width=True,
            hide_index=True
        )

    else:
        st.info("Transporter column not available in the Back Up sheet.")


# ============================================================
# Tab 3: Sales / State / Zone
# ============================================================

with tab3:
    st.markdown(
        '<div class="section-title">Sales Vertical, State, Zone and Plant Analysis</div>',
        unsafe_allow_html=True
    )

    c1, c2 = st.columns(2)

    with c1:
        vertical = (
            filtered.groupby("Sales Verticle", dropna=False)
            .agg(
                Net_Sales=("Net Sales", "sum"),
                Freight=("Freight", "sum"),
                Invoice_Count=("Invoice No", "nunique")
            )
            .reset_index()
            .sort_values("Freight", ascending=False)
        ) if "Sales Verticle" in filtered.columns and "Invoice No" in filtered.columns else pd.DataFrame()

        make_bar(
            vertical,
            "Sales Verticle",
            "Freight",
            "Sales Vertical-wise Freight"
        )

    with c2:
        plant = (
            filtered.groupby("Plant Name", dropna=False)
            .agg(
                Net_Sales=("Net Sales", "sum"),
                Freight=("Freight", "sum"),
                Invoice_Count=("Invoice No", "nunique")
            )
            .reset_index()
            .sort_values("Freight", ascending=False)
        ) if "Plant Name" in filtered.columns and "Invoice No" in filtered.columns else pd.DataFrame()

        make_bar(
            plant,
            "Plant Name",
            "Freight",
            "Plant-wise Freight"
        )

    c3, c4 = st.columns(2)

    with c3:
        bill_type = (
            filtered.groupby("Billing Type", dropna=False)
            .agg(
                Net_Sales=("Net Sales", "sum"),
                Freight=("Freight", "sum")
            )
            .reset_index()
            .sort_values("Freight", ascending=False)
        ) if "Billing Type" in filtered.columns else pd.DataFrame()

        make_bar(
            bill_type,
            "Billing Type",
            "Freight",
            "Billing Type-wise Freight"
        )

    with c4:
        exp_type = (
            filtered.groupby("Expense type", dropna=False)
            .agg(
                Freight=("Freight", "sum"),
                Received_Bill=("Received Bill Amount", "sum")
            )
            .reset_index()
            .sort_values("Freight", ascending=False)
        ) if "Expense type" in filtered.columns else pd.DataFrame()

        if not exp_type.empty:
            exp_plot = exp_type.melt(
                id_vars="Expense type",
                value_vars=["Freight", "Received_Bill"],
                var_name="Metric",
                value_name="Amount"
            )

            make_bar(
                exp_plot,
                "Expense type",
                "Amount",
                "Expense Type-wise Freight vs Received Bill",
                color="Metric"
            )

    if not vertical.empty:
        vertical["Freight %"] = np.where(
            vertical["Net_Sales"] != 0,
            vertical["Freight"] / vertical["Net_Sales"] * 100,
            0
        )

        st.markdown("#### Sales Vertical Summary")
        st.dataframe(
            vertical,
            use_container_width=True,
            hide_index=True
        )


# ============================================================
# Tab 4: C&FA Expense Analysis
# ============================================================

with tab4:
    st.markdown(
        '<div class="section-title">C&FA Expense Analysis</div>',
        unsafe_allow_html=True
    )

    if cfa_filtered.empty:
        st.info("No C&FA expense data available for current filters.")

    else:
        c1, c2 = st.columns(2)

        cfa_by_name = (
            cfa_filtered.groupby("CFA Name", dropna=False)
            .agg(
                Expense_Amount=("Expense Amount", "sum")
            )
            .reset_index()
            .sort_values("Expense_Amount", ascending=False)
        )

        cfa_by_nature = (
            cfa_filtered.groupby("Nature", dropna=False)
            .agg(
                Expense_Amount=("Expense Amount", "sum")
            )
            .reset_index()
            .sort_values("Expense_Amount", ascending=False)
        )

        with c1:
            make_bar(
                cfa_by_name,
                "CFA Name",
                "Expense_Amount",
                "C&FA-wise Expense"
            )

        with c2:
            make_bar(
                cfa_by_nature.head(20),
                "Nature",
                "Expense_Amount",
                "Nature-wise Expense"
            )

        c3, c4 = st.columns(2)

        with c3:
            cfa_month = (
                cfa_filtered.groupby("Month", dropna=False)
                .agg(
                    Expense_Amount=("Expense Amount", "sum")
                )
                .reset_index()
            )

            cfa_month, cfa_month_order = fix_month_order(cfa_month, "Month")

            make_bar(
                cfa_month,
                "Month",
                "Expense_Amount",
                "Month-wise C&FA Expense",
                category_orders={"Month": cfa_month_order}
            )

        with c4:
            cfa_loc = (
                cfa_filtered.groupby("CFA Location", dropna=False)
                .agg(
                    Expense_Amount=("Expense Amount", "sum")
                )
                .reset_index()
                .sort_values("Expense_Amount", ascending=False)
            )

            make_bar(
                cfa_loc,
                "CFA Location",
                "Expense_Amount",
                "Location-wise C&FA Expense"
            )

        st.markdown("#### C&FA Expense Detail")

        st.dataframe(
            cfa_filtered,
            use_container_width=True,
            hide_index=True
        )


# ============================================================
# Tab 5: C&FA Reconciliation
# ============================================================

with tab5:
    st.markdown(
        '<div class="section-title">C&FA vs Back Up Reconciliation</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="small-caption">This compares Back Up sheet transporter/CFA received bill and freight with the C&FA-Exp sheet expense amount using normalized names.</div>',
        unsafe_allow_html=True
    )

    display_recon = recon_df.copy()

    display_recon = display_recon[
        (display_recon["Backup_Freight"] != 0)
        | (display_recon["Backup_Received_Bill"] != 0)
        | (display_recon["CFA_Expense"] != 0)
    ]

    display_recon = display_recon.sort_values(
        "CFA Expense vs Received Bill Variance",
        ascending=False
    )

    c1, c2 = st.columns(2)

    with c1:
        make_bar(
            display_recon.head(20),
            "Back Up Transporter/CFA",
            "CFA Expense vs Received Bill Variance",
            "C&FA Expense vs Received Bill Variance - Top"
        )

    with c2:
        make_bar(
            display_recon.sort_values(
                "CFA_Expense",
                ascending=False
            ).head(20),
            "CFA Name",
            "CFA_Expense",
            "Top C&FA Expenses as per C&FA-Exp"
        )

    st.dataframe(
        display_recon,
        use_container_width=True,
        hide_index=True
    )


# ============================================================
# Tab 6: Data and Download
# ============================================================

with tab6:
    st.markdown(
        '<div class="section-title">Data Preview & Export</div>',
        unsafe_allow_html=True
    )

    st.markdown("#### Filtered Back Up Data")
    st.dataframe(
        filtered,
        use_container_width=True,
        hide_index=True
    )

    st.markdown("#### Filtered C&FA Expense Data")
    st.dataframe(
        cfa_filtered,
        use_container_width=True,
        hide_index=True
    )

    download_excel_button(
        {
            "Filtered Back Up": filtered,
            "Filtered CFA Expense": cfa_filtered,
            "CFA Reconciliation": recon_df,
        },
        "Freight_Analysis_Filtered_Report.xlsx",
    )


# ============================================================
# Footer
# ============================================================

st.caption(
    "Dashboard logic: Back Up sheet = freight, billing, transporter and sales analysis. "
    "C&FA-Exp sheet = C&FA name/location/nature-wise expense analysis. "
    "Net Sales is calculated after reducing Sales Return / Credit Note values."
)
