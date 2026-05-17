import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from io import BytesIO

# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="Freight Analysis Dashboard",
    page_icon="🚚",
    layout="wide"
)

# =========================================================
# CUSTOM CSS
# =========================================================

st.markdown(
    """
    <style>
        .main {
            background-color: #f7f9fc;
        }

        .main-title {
            font-size: 34px;
            font-weight: 800;
            color: #111827;
            margin-bottom: 0px;
        }

        .sub-title {
            font-size: 15px;
            color: #6b7280;
            margin-top: 0px;
            margin-bottom: 22px;
        }

        .section-title {
            font-size: 24px;
            font-weight: 750;
            color: #111827;
            margin-top: 25px;
            margin-bottom: 10px;
        }

        .metric-card {
            background: #ffffff;
            padding: 18px 20px;
            border-radius: 18px;
            box-shadow: 0 8px 20px rgba(15, 23, 42, 0.08);
            border: 1px solid #e5e7eb;
            min-height: 115px;
        }

        .metric-label {
            font-size: 13px;
            color: #6b7280;
            font-weight: 600;
            margin-bottom: 8px;
        }

        .metric-value {
            font-size: 25px;
            color: #111827;
            font-weight: 800;
        }

        .metric-small {
            font-size: 12px;
            color: #9ca3af;
            margin-top: 6px;
        }

        div[data-testid="stMetricValue"] {
            font-size: 24px;
            font-weight: 800;
        }

        .stTabs [data-baseweb="tab-list"] {
            gap: 8px;
        }

        .stTabs [data-baseweb="tab"] {
            height: 45px;
            border-radius: 12px;
            padding: 8px 18px;
            background-color: #ffffff;
            border: 1px solid #e5e7eb;
            font-weight: 700;
        }

        .stTabs [aria-selected="true"] {
            background-color: #0f172a !important;
            color: white !important;
        }
    </style>
    """,
    unsafe_allow_html=True
)

# =========================================================
# HELPER FUNCTIONS
# =========================================================

def clean_col_name(col):
    col = str(col).strip()
    col = col.replace("\n", " ")
    col = col.replace("\r", " ")
    col = " ".join(col.split())
    col = col.replace(" ", "_")
    col = col.replace("-", "_")
    col = col.replace("/", "_")
    col = col.replace(".", "")
    col = col.replace("&", "and")
    col = col.replace("%", "Percent")
    return col


def clean_dataframe_columns(df):
    df = df.copy()
    df.columns = [clean_col_name(c) for c in df.columns]
    return df


def to_number(series):
    return (
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("₹", "", regex=False)
        .str.replace("(", "-", regex=False)
        .str.replace(")", "", regex=False)
        .str.strip()
        .replace(["", "nan", "None", "NaT"], np.nan)
        .pipe(pd.to_numeric, errors="coerce")
        .fillna(0)
    )


def format_inr(value):
    try:
        value = float(value)
    except Exception:
        value = 0

    negative = value < 0
    value = abs(value)

    if value >= 10_000_000:
        result = f"₹{value / 10_000_000:,.2f} Cr"
    elif value >= 100_000:
        result = f"₹{value / 100_000:,.2f} L"
    elif value >= 1_000:
        result = f"₹{value / 1_000:,.2f} K"
    else:
        result = f"₹{value:,.0f}"

    return f"-{result}" if negative else result


def safe_find_column(df, possible_names):
    cols = list(df.columns)

    normalized_cols = {
        c.lower().replace("_", "").replace(" ", "").replace(".", "").replace("-", "").replace("&", "and"): c
        for c in cols
    }

    for name in possible_names:
        key = name.lower().replace("_", "").replace(" ", "").replace(".", "").replace("-", "").replace("&", "and")
        if key in normalized_cols:
            return normalized_cols[key]

    for c in cols:
        c_norm = c.lower().replace("_", " ")
        for name in possible_names:
            n_norm = name.lower().replace("_", " ")
            if n_norm in c_norm:
                return c

    return None


def make_month_from_date(df):
    df = df.copy()

    possible_date_cols = [
        "Date",
        "Billing_Date",
        "Bill_Date",
        "Invoice_Date",
        "Month",
        "Posting_Date",
        "Created_on",
        "Created_On"
    ]

    month_col = safe_find_column(df, ["Month", "Months"])

    if month_col:
        df["Month"] = df[month_col].astype(str).str.strip()
        return df

    date_col = safe_find_column(df, possible_date_cols)

    if date_col:
        date_data = pd.to_datetime(df[date_col], errors="coerce")
        df["Month"] = date_data.dt.strftime("%b-%y")
    else:
        df["Month"] = "Unknown"

    return df


def fix_month_order(df, month_col="Month"):
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
            errors="coerce"
        )

    if df["_Month_Date"].isna().all():
        month_map = {
            "jan": 1, "january": 1,
            "feb": 2, "february": 2,
            "mar": 3, "march": 3,
            "apr": 4, "april": 4,
            "may": 5,
            "jun": 6, "june": 6,
            "jul": 7, "july": 7,
            "aug": 8, "august": 8,
            "sep": 9, "sept": 9, "september": 9,
            "oct": 10, "october": 10,
            "nov": 11, "november": 11,
            "dec": 12, "december": 12
        }

        temp_month = month_text.str.lower().str[:3].map(month_map)
        df["_Month_Date"] = pd.to_datetime(
            {
                "year": 2026,
                "month": temp_month.fillna(12),
                "day": 1
            },
            errors="coerce"
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


def get_top_table(df, group_col, value_cols, top_n=20):
    if not group_col or group_col not in df.columns:
        return pd.DataFrame()

    available_values = [c for c in value_cols if c in df.columns]

    if not available_values:
        return pd.DataFrame()

    out = (
        df.groupby(group_col, dropna=False)[available_values]
        .sum()
        .reset_index()
        .sort_values(available_values[0], ascending=False)
        .head(top_n)
    )

    return out


def create_download_excel(dataframes_dict):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for sheet_name, data in dataframes_dict.items():
            safe_sheet = str(sheet_name)[:31]
            data.to_excel(writer, sheet_name=safe_sheet, index=False)

    output.seek(0)
    return output


def show_metric_card(label, value, small_text=""):
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">{label}</div>
            <div class="metric-value">{value}</div>
            <div class="metric-small">{small_text}</div>
        </div>
        """,
        unsafe_allow_html=True
    )


# =========================================================
# FILE READING
# =========================================================

@st.cache_data(show_spinner=False)
def read_uploaded_file(uploaded_file):
    suffix = uploaded_file.name.split(".")[-1].lower()

    if suffix == "xlsb":
        backup_df = pd.read_excel(
            uploaded_file,
            sheet_name="Back Up",
            engine="pyxlsb"
        )

        cfa_df = pd.read_excel(
            uploaded_file,
            sheet_name="C&FA-Exp",
            engine="pyxlsb"
        )

    else:
        backup_df = pd.read_excel(
            uploaded_file,
            sheet_name="Back Up"
        )

        cfa_df = pd.read_excel(
            uploaded_file,
            sheet_name="C&FA-Exp"
        )

    backup_df = clean_dataframe_columns(backup_df)
    cfa_df = clean_dataframe_columns(cfa_df)

    return backup_df, cfa_df


# =========================================================
# DATA PREPARATION
# =========================================================

def prepare_backup_data(df):
    df = df.copy()
    df = make_month_from_date(df)

    net_bill_col = safe_find_column(
        df,
        [
            "Net Bill",
            "Net_Bill",
            "NetBilling",
            "Net Billing",
            "Net Sales",
            "Net_Sales",
            "Billing",
            "Bill Amount"
        ]
    )

    freight_col = safe_find_column(
        df,
        [
            "Freight",
            "Freight Amount",
            "Freight_Amount",
            "Transporter Freight",
            "Transporter_Freight",
            "Freight Expenses",
            "Freight_Expenses"
        ]
    )

    received_bill_col = safe_find_column(
        df,
        [
            "Received Bill",
            "Received_Bill",
            "Bill Received",
            "Bill_Received",
            "Transporter Bill",
            "Transporter_Bill"
        ]
    )

    loading_col = safe_find_column(
        df,
        [
            "Loading",
            "Loading Charges",
            "Loading_Charges",
            "Loading Cost"
        ]
    )

    unloading_col = safe_find_column(
        df,
        [
            "Unloading",
            "Unloading Charges",
            "Unloading_Charges",
            "Unloading Cost"
        ]
    )

    transporter_col = safe_find_column(
        df,
        [
            "Transporter Name",
            "Transporter_Name",
            "Transporter",
            "Name Transporter",
            "Name_Transporter"
        ]
    )

    vertical_col = safe_find_column(
        df,
        [
            "Sales Verticle",
            "Sales Vertical",
            "Sales_Verticle",
            "Sales_Vertical",
            "Vertical"
        ]
    )

    state_col = safe_find_column(
        df,
        [
            "State",
            "Customer State",
            "Customer_State",
            "Ship State",
            "Ship_State"
        ]
    )

    zone_col = safe_find_column(
        df,
        [
            "Zone",
            "Region",
            "Sales Zone",
            "Sales_Zone"
        ]
    )

    plant_col = safe_find_column(
        df,
        [
            "Plant",
            "Plant Code",
            "Plant_Code",
            "Warehouse",
            "WH"
        ]
    )

    city_col = safe_find_column(
        df,
        [
            "City",
            "Customer City",
            "Ship City",
            "Destination City"
        ]
    )

    rename_map = {}

    if net_bill_col:
        rename_map[net_bill_col] = "Net_Bill"
    if freight_col:
        rename_map[freight_col] = "Freight"
    if received_bill_col:
        rename_map[received_bill_col] = "Received_Bill"
    if loading_col:
        rename_map[loading_col] = "Loading"
    if unloading_col:
        rename_map[unloading_col] = "Unloading"
    if transporter_col:
        rename_map[transporter_col] = "Transporter_Name"
    if vertical_col:
        rename_map[vertical_col] = "Sales_Vertical"
    if state_col:
        rename_map[state_col] = "State"
    if zone_col:
        rename_map[zone_col] = "Zone"
    if plant_col:
        rename_map[plant_col] = "Plant"
    if city_col:
        rename_map[city_col] = "City"

    df = df.rename(columns=rename_map)

    for col in ["Net_Bill", "Freight", "Received_Bill", "Loading", "Unloading"]:
        if col not in df.columns:
            df[col] = 0
        df[col] = to_number(df[col])

    for col in ["Transporter_Name", "Sales_Vertical", "State", "Zone", "Plant", "City"]:
        if col not in df.columns:
            df[col] = "Not Available"
        df[col] = df[col].astype(str).str.strip().replace(["", "nan", "None"], "Not Available")

    df["Bill_Variance"] = df["Received_Bill"] - df["Freight"]

    df["Freight_Percent"] = np.where(
        df["Net_Bill"] != 0,
        df["Freight"] / df["Net_Bill"] * 100,
        0
    )

    df["Loading_Unloading"] = df["Loading"] + df["Unloading"]

    df["Total_Logistic_Cost"] = (
        df["Freight"]
        + df["Loading"]
        + df["Unloading"]
    )

    df["Logistic_Cost_Percent"] = np.where(
        df["Net_Bill"] != 0,
        df["Total_Logistic_Cost"] / df["Net_Bill"] * 100,
        0
    )

    return df


def prepare_cfa_data(df):
    df = df.copy()
    df = make_month_from_date(df)

    cfa_col = safe_find_column(
        df,
        [
            "CFA",
            "C&FA",
            "C and FA",
            "CFA Name",
            "CFA_Name",
            "C&FA Name",
            "C&FA_Name",
            "CFA Wise Names",
            "CFA_Wise_Names",
            "Name"
        ]
    )

    nature_col = safe_find_column(
        df,
        [
            "Nature of Expense",
            "Nature_of_Expense",
            "Expense Nature",
            "Expense_Nature",
            "Nature",
            "Expense Type",
            "Expense_Type"
        ]
    )

    amount_col = safe_find_column(
        df,
        [
            "Amount",
            "Expense Amount",
            "Expense_Amount",
            "Value",
            "Total",
            "Total Amount",
            "Total_Amount"
        ]
    )

    location_col = safe_find_column(
        df,
        [
            "Location",
            "City",
            "Place",
            "Branch",
            "Depot"
        ]
    )

    rename_map = {}

    if cfa_col:
        rename_map[cfa_col] = "CFA_Name"
    if nature_col:
        rename_map[nature_col] = "Nature_of_Expense"
    if amount_col:
        rename_map[amount_col] = "Expense_Amount"
    if location_col:
        rename_map[location_col] = "Location"

    df = df.rename(columns=rename_map)

    if "CFA_Name" not in df.columns:
        df["CFA_Name"] = "Not Available"

    if "Nature_of_Expense" not in df.columns:
        df["Nature_of_Expense"] = "Not Available"

    if "Expense_Amount" not in df.columns:
        df["Expense_Amount"] = 0

    if "Location" not in df.columns:
        df["Location"] = "Not Available"

    df["CFA_Name"] = df["CFA_Name"].astype(str).str.strip().replace(["", "nan", "None"], "Not Available")
    df["Nature_of_Expense"] = df["Nature_of_Expense"].astype(str).str.strip().replace(["", "nan", "None"], "Not Available")
    df["Location"] = df["Location"].astype(str).str.strip().replace(["", "nan", "None"], "Not Available")
    df["Expense_Amount"] = to_number(df["Expense_Amount"])

    return df


# =========================================================
# HEADER
# =========================================================

st.markdown('<div class="main-title">🚚 Freight Analysis Dashboard</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="sub-title">Back Up Sheet vs C&FA Expense Sheet | Net Bill, Freight, Received Bill, Transporter, Sales Vertical, State, Zone and C&FA Expense Analysis</div>',
    unsafe_allow_html=True
)

# =========================================================
# UPLOAD FILE
# =========================================================

uploaded_file = st.file_uploader(
    "Upload Freight Excel File",
    type=["xlsb", "xlsx", "xls"]
)

if uploaded_file is None:
    st.info("Please upload your freight file containing sheets named 'Back Up' and 'C&FA-Exp'.")
    st.stop()

try:
    raw_backup_df, raw_cfa_df = read_uploaded_file(uploaded_file)
except Exception as e:
    st.error("Unable to read the uploaded file. Please check sheet names: 'Back Up' and 'C&FA-Exp'.")
    st.exception(e)
    st.stop()

backup_df = prepare_backup_data(raw_backup_df)
cfa_df = prepare_cfa_data(raw_cfa_df)

# =========================================================
# SIDEBAR FILTERS
# =========================================================

st.sidebar.header("🔎 Dashboard Filters")

month_list_df, month_order_all = fix_month_order(
    backup_df[["Month"]].drop_duplicates(),
    "Month"
)

month_options = month_order_all

selected_months = st.sidebar.multiselect(
    "Select Month",
    options=month_options,
    default=month_options
)

selected_verticals = st.sidebar.multiselect(
    "Select Sales Vertical",
    options=sorted(backup_df["Sales_Vertical"].dropna().unique().tolist()),
    default=sorted(backup_df["Sales_Vertical"].dropna().unique().tolist())
)

selected_states = st.sidebar.multiselect(
    "Select State",
    options=sorted(backup_df["State"].dropna().unique().tolist()),
    default=sorted(backup_df["State"].dropna().unique().tolist())
)

selected_zones = st.sidebar.multiselect(
    "Select Zone",
    options=sorted(backup_df["Zone"].dropna().unique().tolist()),
    default=sorted(backup_df["Zone"].dropna().unique().tolist())
)

selected_transporters = st.sidebar.multiselect(
    "Select Transporter",
    options=sorted(backup_df["Transporter_Name"].dropna().unique().tolist()),
    default=sorted(backup_df["Transporter_Name"].dropna().unique().tolist())
)

filtered_backup_df = backup_df[
    backup_df["Month"].isin(selected_months)
    & backup_df["Sales_Vertical"].isin(selected_verticals)
    & backup_df["State"].isin(selected_states)
    & backup_df["Zone"].isin(selected_zones)
    & backup_df["Transporter_Name"].isin(selected_transporters)
].copy()

filtered_cfa_df = cfa_df[
    cfa_df["Month"].isin(selected_months)
].copy()

# =========================================================
# KPI CALCULATION
# =========================================================

total_net_bill = filtered_backup_df["Net_Bill"].sum()
total_freight = filtered_backup_df["Freight"].sum()
total_received_bill = filtered_backup_df["Received_Bill"].sum()
total_variance = filtered_backup_df["Bill_Variance"].sum()
total_loading_unloading = filtered_backup_df["Loading_Unloading"].sum()
total_logistic_cost = filtered_backup_df["Total_Logistic_Cost"].sum()
total_cfa_expense = filtered_cfa_df["Expense_Amount"].sum()

freight_percent = (total_freight / total_net_bill * 100) if total_net_bill else 0
logistic_percent = (total_logistic_cost / total_net_bill * 100) if total_net_bill else 0
cfa_percent = (total_cfa_expense / total_net_bill * 100) if total_net_bill else 0

# =========================================================
# KPI CARDS
# =========================================================

st.markdown('<div class="section-title">📌 Freight Summary</div>', unsafe_allow_html=True)

kpi1, kpi2, kpi3, kpi4 = st.columns(4)

with kpi1:
    show_metric_card("Net Bill", format_inr(total_net_bill), "Total billing value")

with kpi2:
    show_metric_card("Freight Expense", format_inr(total_freight), f"{freight_percent:.2f}% of Net Bill")

with kpi3:
    show_metric_card("Received Bill", format_inr(total_received_bill), "Transporter bill received")

with kpi4:
    show_metric_card("Bill Variance", format_inr(total_variance), "Received Bill - Freight")

kpi5, kpi6, kpi7, kpi8 = st.columns(4)

with kpi5:
    show_metric_card("Loading + Unloading", format_inr(total_loading_unloading), "Additional logistic cost")

with kpi6:
    show_metric_card("Total Logistic Cost", format_inr(total_logistic_cost), f"{logistic_percent:.2f}% of Net Bill")

with kpi7:
    show_metric_card("C&FA Expense", format_inr(total_cfa_expense), f"{cfa_percent:.2f}% of Net Bill")

with kpi8:
    show_metric_card("Total Records", f"{len(filtered_backup_df):,}", "Back Up sheet records")

# =========================================================
# TABS
# =========================================================

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
    [
        "📊 Freight Overview",
        "🚚 Transporter Analysis",
        "🏢 Sales Vertical / State / Zone",
        "🏬 C&FA Expense",
        "🔁 Reconciliation",
        "📄 Data & Download"
    ]
)

# =========================================================
# TAB 1 - FREIGHT OVERVIEW
# =========================================================

with tab1:
    st.markdown('<div class="section-title">Freight Overview</div>', unsafe_allow_html=True)

    month_summary = (
        filtered_backup_df
        .groupby("Month", dropna=False)
        .agg(
            Net_Bill=("Net_Bill", "sum"),
            Freight=("Freight", "sum"),
            Received_Bill=("Received_Bill", "sum"),
            Bill_Variance=("Bill_Variance", "sum"),
            Loading_Unloading=("Loading_Unloading", "sum"),
            Total_Logistic_Cost=("Total_Logistic_Cost", "sum")
        )
        .reset_index()
    )

    month_summary["Freight_Percent"] = np.where(
        month_summary["Net_Bill"] != 0,
        month_summary["Freight"] / month_summary["Net_Bill"] * 100,
        0
    )

    month_summary["Logistic_Cost_Percent"] = np.where(
        month_summary["Net_Bill"] != 0,
        month_summary["Total_Logistic_Cost"] / month_summary["Net_Bill"] * 100,
        0
    )

    month_summary, month_order = fix_month_order(month_summary, "Month")

    col1, col2 = st.columns(2)

    with col1:
        fig = px.bar(
            month_summary,
            x="Month",
            y=["Net_Bill", "Freight", "Received_Bill"],
            barmode="group",
            title="Month-wise Net Bill vs Freight vs Received Bill",
            text_auto=".3s",
            category_orders={"Month": month_order}
        )

        fig.update_layout(
            xaxis_title="Month",
            yaxis_title="Amount",
            legend_title_text="Particulars"
        )

        st.plotly_chart(fig, use_container_width=True)

    with col2:
        fig = px.line(
            month_summary,
            x="Month",
            y="Freight_Percent",
            markers=True,
            title="Month-wise Freight % on Net Bill",
            category_orders={"Month": month_order}
        )

        fig.update_layout(
            xaxis_title="Month",
            yaxis_title="Freight %"
        )

        st.plotly_chart(fig, use_container_width=True)

    col3, col4 = st.columns(2)

    with col3:
        fig = px.bar(
            month_summary,
            x="Month",
            y=["Loading_Unloading", "Total_Logistic_Cost"],
            barmode="group",
            title="Month-wise Loading/Unloading and Total Logistic Cost",
            text_auto=".3s",
            category_orders={"Month": month_order}
        )

        fig.update_layout(
            xaxis_title="Month",
            yaxis_title="Amount",
            legend_title_text="Particulars"
        )

        st.plotly_chart(fig, use_container_width=True)

    with col4:
        fig = px.line(
            month_summary,
            x="Month",
            y="Logistic_Cost_Percent",
            markers=True,
            title="Month-wise Total Logistic Cost % on Net Bill",
            category_orders={"Month": month_order}
        )

        fig.update_layout(
            xaxis_title="Month",
            yaxis_title="Total Logistic Cost %"
        )

        st.plotly_chart(fig, use_container_width=True)

    st.dataframe(month_summary.drop(columns=["_Month_Date"], errors="ignore"), use_container_width=True)

# =========================================================
# TAB 2 - TRANSPORTER ANALYSIS
# =========================================================

with tab2:
    st.markdown('<div class="section-title">Transporter Wise Analysis</div>', unsafe_allow_html=True)

    transporter_summary = (
        filtered_backup_df
        .groupby("Transporter_Name", dropna=False)
        .agg(
            Net_Bill=("Net_Bill", "sum"),
            Freight=("Freight", "sum"),
            Received_Bill=("Received_Bill", "sum"),
            Bill_Variance=("Bill_Variance", "sum"),
            Loading_Unloading=("Loading_Unloading", "sum"),
            Total_Logistic_Cost=("Total_Logistic_Cost", "sum")
        )
        .reset_index()
    )

    transporter_summary["Freight_Percent"] = np.where(
        transporter_summary["Net_Bill"] != 0,
        transporter_summary["Freight"] / transporter_summary["Net_Bill"] * 100,
        0
    )

    transporter_summary = transporter_summary.sort_values("Freight", ascending=False)

    col1, col2 = st.columns(2)

    with col1:
        top_transporters = transporter_summary.head(20)

        fig = px.bar(
            top_transporters,
            x="Transporter_Name",
            y="Freight",
            title="Top Transporters by Freight Expense",
            text_auto=".3s"
        )

        fig.update_layout(
            xaxis_title="Transporter",
            yaxis_title="Freight",
            xaxis_tickangle=-45
        )

        st.plotly_chart(fig, use_container_width=True)

    with col2:
        top_variance = transporter_summary.sort_values("Bill_Variance", ascending=False).head(20)

        fig = px.bar(
            top_variance,
            x="Transporter_Name",
            y="Bill_Variance",
            title="Top Transporters by Bill Variance",
            text_auto=".3s"
        )

        fig.update_layout(
            xaxis_title="Transporter",
            yaxis_title="Variance",
            xaxis_tickangle=-45
        )

        st.plotly_chart(fig, use_container_width=True)

    col3, col4 = st.columns(2)

    with col3:
        fig = px.bar(
            transporter_summary.head(20),
            x="Transporter_Name",
            y=["Freight", "Received_Bill"],
            barmode="group",
            title="Transporter Freight vs Received Bill",
            text_auto=".3s"
        )

        fig.update_layout(
            xaxis_title="Transporter",
            yaxis_title="Amount",
            xaxis_tickangle=-45
        )

        st.plotly_chart(fig, use_container_width=True)

    with col4:
        fig = px.bar(
            transporter_summary.head(20),
            x="Transporter_Name",
            y="Freight_Percent",
            title="Transporter Wise Freight % on Net Bill",
            text_auto=".2f"
        )

        fig.update_layout(
            xaxis_title="Transporter",
            yaxis_title="Freight %",
            xaxis_tickangle=-45
        )

        st.plotly_chart(fig, use_container_width=True)

    st.dataframe(transporter_summary, use_container_width=True)

# =========================================================
# TAB 3 - SALES VERTICAL / STATE / ZONE
# =========================================================

with tab3:
    st.markdown('<div class="section-title">Sales Vertical, State and Zone Analysis</div>', unsafe_allow_html=True)

    vertical_summary = (
        filtered_backup_df
        .groupby("Sales_Vertical", dropna=False)
        .agg(
            Net_Bill=("Net_Bill", "sum"),
            Freight=("Freight", "sum"),
            Received_Bill=("Received_Bill", "sum"),
            Total_Logistic_Cost=("Total_Logistic_Cost", "sum")
        )
        .reset_index()
    )

    vertical_summary["Freight_Percent"] = np.where(
        vertical_summary["Net_Bill"] != 0,
        vertical_summary["Freight"] / vertical_summary["Net_Bill"] * 100,
        0
    )

    vertical_summary = vertical_summary.sort_values("Freight", ascending=False)

    state_summary = (
        filtered_backup_df
        .groupby("State", dropna=False)
        .agg(
            Net_Bill=("Net_Bill", "sum"),
            Freight=("Freight", "sum"),
            Received_Bill=("Received_Bill", "sum"),
            Total_Logistic_Cost=("Total_Logistic_Cost", "sum")
        )
        .reset_index()
    )

    state_summary["Freight_Percent"] = np.where(
        state_summary["Net_Bill"] != 0,
        state_summary["Freight"] / state_summary["Net_Bill"] * 100,
        0
    )

    state_summary = state_summary.sort_values("Freight", ascending=False)

    zone_summary = (
        filtered_backup_df
        .groupby("Zone", dropna=False)
        .agg(
            Net_Bill=("Net_Bill", "sum"),
            Freight=("Freight", "sum"),
            Received_Bill=("Received_Bill", "sum"),
            Total_Logistic_Cost=("Total_Logistic_Cost", "sum")
        )
        .reset_index()
    )

    zone_summary["Freight_Percent"] = np.where(
        zone_summary["Net_Bill"] != 0,
        zone_summary["Freight"] / zone_summary["Net_Bill"] * 100,
        0
    )

    zone_summary = zone_summary.sort_values("Freight", ascending=False)

    col1, col2 = st.columns(2)

    with col1:
        fig = px.bar(
            vertical_summary,
            x="Sales_Vertical",
            y=["Net_Bill", "Freight"],
            barmode="group",
            title="Sales Vertical Wise Net Bill vs Freight",
            text_auto=".3s"
        )

        fig.update_layout(
            xaxis_title="Sales Vertical",
            yaxis_title="Amount",
            xaxis_tickangle=-45
        )

        st.plotly_chart(fig, use_container_width=True)

    with col2:
        fig = px.pie(
            vertical_summary,
            names="Sales_Vertical",
            values="Freight",
            title="Sales Vertical Wise Freight Share"
        )

        st.plotly_chart(fig, use_container_width=True)

    col3, col4 = st.columns(2)

    with col3:
        fig = px.bar(
            state_summary.head(20),
            x="State",
            y="Freight",
            title="Top State Wise Freight",
            text_auto=".3s"
        )

        fig.update_layout(
            xaxis_title="State",
            yaxis_title="Freight",
            xaxis_tickangle=-45
        )

        st.plotly_chart(fig, use_container_width=True)

    with col4:
        fig = px.bar(
            zone_summary,
            x="Zone",
            y="Freight",
            title="Zone Wise Freight",
            text_auto=".3s"
        )

        fig.update_layout(
            xaxis_title="Zone",
            yaxis_title="Freight"
        )

        st.plotly_chart(fig, use_container_width=True)

    st.subheader("Sales Vertical Summary")
    st.dataframe(vertical_summary, use_container_width=True)

    st.subheader("State Summary")
    st.dataframe(state_summary, use_container_width=True)

    st.subheader("Zone Summary")
    st.dataframe(zone_summary, use_container_width=True)

# =========================================================
# TAB 4 - CFA EXPENSE
# =========================================================

with tab4:
    st.markdown('<div class="section-title">C&FA Expense Analysis</div>', unsafe_allow_html=True)

    cfa_month_summary = (
        filtered_cfa_df
        .groupby("Month", dropna=False)
        .agg(
            Expense_Amount=("Expense_Amount", "sum")
        )
        .reset_index()
    )

    cfa_month_summary, cfa_month_order = fix_month_order(cfa_month_summary, "Month")

    cfa_summary = (
        filtered_cfa_df
        .groupby("CFA_Name", dropna=False)
        .agg(
            Expense_Amount=("Expense_Amount", "sum")
        )
        .reset_index()
        .sort_values("Expense_Amount", ascending=False)
    )

    nature_summary = (
        filtered_cfa_df
        .groupby("Nature_of_Expense", dropna=False)
        .agg(
            Expense_Amount=("Expense_Amount", "sum")
        )
        .reset_index()
        .sort_values("Expense_Amount", ascending=False)
    )

    location_summary = (
        filtered_cfa_df
        .groupby("Location", dropna=False)
        .agg(
            Expense_Amount=("Expense_Amount", "sum")
        )
        .reset_index()
        .sort_values("Expense_Amount", ascending=False)
    )

    col1, col2 = st.columns(2)

    with col1:
        fig = px.bar(
            cfa_month_summary,
            x="Month",
            y="Expense_Amount",
            title="Month-wise C&FA Expense",
            text_auto=".3s",
            category_orders={"Month": cfa_month_order}
        )

        fig.update_layout(
            xaxis_title="Month",
            yaxis_title="Expense Amount"
        )

        st.plotly_chart(fig, use_container_width=True)

    with col2:
        fig = px.pie(
            nature_summary,
            names="Nature_of_Expense",
            values="Expense_Amount",
            title="Nature of Expense Wise C&FA Expense"
        )

        st.plotly_chart(fig, use_container_width=True)

    col3, col4 = st.columns(2)

    with col3:
        fig = px.bar(
            cfa_summary.head(20),
            x="CFA_Name",
            y="Expense_Amount",
            title="Top C&FA Wise Expense",
            text_auto=".3s"
        )

        fig.update_layout(
            xaxis_title="C&FA",
            yaxis_title="Expense Amount",
            xaxis_tickangle=-45
        )

        st.plotly_chart(fig, use_container_width=True)

    with col4:
        fig = px.bar(
            location_summary.head(20),
            x="Location",
            y="Expense_Amount",
            title="Location Wise C&FA Expense",
            text_auto=".3s"
        )

        fig.update_layout(
            xaxis_title="Location",
            yaxis_title="Expense Amount",
            xaxis_tickangle=-45
        )

        st.plotly_chart(fig, use_container_width=True)

    st.subheader("C&FA Wise Summary")
    st.dataframe(cfa_summary, use_container_width=True)

    st.subheader("Nature of Expense Summary")
    st.dataframe(nature_summary, use_container_width=True)

    st.subheader("Location Wise Summary")
    st.dataframe(location_summary, use_container_width=True)

# =========================================================
# TAB 5 - RECONCILIATION
# =========================================================

with tab5:
    st.markdown('<div class="section-title">Back Up vs C&FA Expense Reconciliation</div>', unsafe_allow_html=True)

    backup_month = (
        filtered_backup_df
        .groupby("Month", dropna=False)
        .agg(
            Net_Bill=("Net_Bill", "sum"),
            Freight=("Freight", "sum"),
            Received_Bill=("Received_Bill", "sum"),
            Loading_Unloading=("Loading_Unloading", "sum"),
            Total_Logistic_Cost=("Total_Logistic_Cost", "sum")
        )
        .reset_index()
    )

    cfa_month = (
        filtered_cfa_df
        .groupby("Month", dropna=False)
        .agg(
            CFA_Expense=("Expense_Amount", "sum")
        )
        .reset_index()
    )

    recon_month = backup_month.merge(
        cfa_month,
        on="Month",
        how="outer"
    ).fillna(0)

    recon_month["Total_Freight_Plus_CFA"] = recon_month["Total_Logistic_Cost"] + recon_month["CFA_Expense"]

    recon_month["CFA_Percent_on_Net_Bill"] = np.where(
        recon_month["Net_Bill"] != 0,
        recon_month["CFA_Expense"] / recon_month["Net_Bill"] * 100,
        0
    )

    recon_month["Total_Logistic_Plus_CFA_Percent"] = np.where(
        recon_month["Net_Bill"] != 0,
        recon_month["Total_Freight_Plus_CFA"] / recon_month["Net_Bill"] * 100,
        0
    )

    recon_month, recon_month_order = fix_month_order(recon_month, "Month")

    col1, col2 = st.columns(2)

    with col1:
        fig = px.bar(
            recon_month,
            x="Month",
            y=["Total_Logistic_Cost", "CFA_Expense", "Total_Freight_Plus_CFA"],
            barmode="group",
            title="Month-wise Logistic Cost vs C&FA Expense",
            text_auto=".3s",
            category_orders={"Month": recon_month_order}
        )

        fig.update_layout(
            xaxis_title="Month",
            yaxis_title="Amount",
            legend_title_text="Particulars"
        )

        st.plotly_chart(fig, use_container_width=True)

    with col2:
        fig = px.line(
            recon_month,
            x="Month",
            y=["CFA_Percent_on_Net_Bill", "Total_Logistic_Plus_CFA_Percent"],
            markers=True,
            title="C&FA and Total Logistic + C&FA % on Net Bill",
            category_orders={"Month": recon_month_order}
        )

        fig.update_layout(
            xaxis_title="Month",
            yaxis_title="Percentage"
        )

        st.plotly_chart(fig, use_container_width=True)

    st.dataframe(recon_month.drop(columns=["_Month_Date"], errors="ignore"), use_container_width=True)

# =========================================================
# TAB 6 - DATA & DOWNLOAD
# =========================================================

with tab6:
    st.markdown('<div class="section-title">Data Preview and Download</div>', unsafe_allow_html=True)

    st.subheader("Filtered Back Up Data")
    st.dataframe(filtered_backup_df, use_container_width=True)

    st.subheader("Filtered C&FA Expense Data")
    st.dataframe(filtered_cfa_df, use_container_width=True)

    month_export = (
        filtered_backup_df
        .groupby("Month", dropna=False)
        .agg(
            Net_Bill=("Net_Bill", "sum"),
            Freight=("Freight", "sum"),
            Received_Bill=("Received_Bill", "sum"),
            Bill_Variance=("Bill_Variance", "sum"),
            Loading_Unloading=("Loading_Unloading", "sum"),
            Total_Logistic_Cost=("Total_Logistic_Cost", "sum")
        )
        .reset_index()
    )

    month_export["Freight_Percent"] = np.where(
        month_export["Net_Bill"] != 0,
        month_export["Freight"] / month_export["Net_Bill"] * 100,
        0
    )

    month_export, _ = fix_month_order(month_export, "Month")
    month_export = month_export.drop(columns=["_Month_Date"], errors="ignore")

    transporter_export = (
        filtered_backup_df
        .groupby("Transporter_Name", dropna=False)
        .agg(
            Net_Bill=("Net_Bill", "sum"),
            Freight=("Freight", "sum"),
            Received_Bill=("Received_Bill", "sum"),
            Bill_Variance=("Bill_Variance", "sum"),
            Total_Logistic_Cost=("Total_Logistic_Cost", "sum")
        )
        .reset_index()
        .sort_values("Freight", ascending=False)
    )

    vertical_export = (
        filtered_backup_df
        .groupby("Sales_Vertical", dropna=False)
        .agg(
            Net_Bill=("Net_Bill", "sum"),
            Freight=("Freight", "sum"),
            Received_Bill=("Received_Bill", "sum"),
            Total_Logistic_Cost=("Total_Logistic_Cost", "sum")
        )
        .reset_index()
        .sort_values("Freight", ascending=False)
    )

    state_export = (
        filtered_backup_df
        .groupby("State", dropna=False)
        .agg(
            Net_Bill=("Net_Bill", "sum"),
            Freight=("Freight", "sum"),
            Received_Bill=("Received_Bill", "sum"),
            Total_Logistic_Cost=("Total_Logistic_Cost", "sum")
        )
        .reset_index()
        .sort_values("Freight", ascending=False)
    )

    cfa_export = (
        filtered_cfa_df
        .groupby("CFA_Name", dropna=False)
        .agg(
            Expense_Amount=("Expense_Amount", "sum")
        )
        .reset_index()
        .sort_values("Expense_Amount", ascending=False)
    )

    nature_export = (
        filtered_cfa_df
        .groupby("Nature_of_Expense", dropna=False)
        .agg(
            Expense_Amount=("Expense_Amount", "sum")
        )
        .reset_index()
        .sort_values("Expense_Amount", ascending=False)
    )

    export_file = create_download_excel(
        {
            "Month Summary": month_export,
            "Transporter Summary": transporter_export,
            "Vertical Summary": vertical_export,
            "State Summary": state_export,
            "CFA Summary": cfa_export,
            "Nature Expense Summary": nature_export,
            "Filtered Back Up": filtered_backup_df,
            "Filtered CFA Expense": filtered_cfa_df
        }
    )

    st.download_button(
        label="📥 Download Freight Analysis Report",
        data=export_file,
        file_name="Freight_Analysis_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# =========================================================
# FOOTER
# =========================================================

st.markdown("---")
st.caption("Freight Analysis Dashboard | Developed for detailed Back Up and C&FA Expense Analysis")
