from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

# ============================================================
# PAGE CONFIGURATION AND DATASET SETTINGS
# ============================================================
st.set_page_config(
    page_title="Sales Intelligence | Analytics Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

BASE_DIR = Path(__file__).resolve().parent
DATA_FILE = BASE_DIR / "kenyan_ecommerce_sales.xlsx"
SHEET_NAME = "sales_2020-10_2023-10(1)"

REQUIRED_COLUMNS = [
    "Order ID", "Sale ID", "Date", "Order", "Transaction type",
    "Sale type", "Sales channel", "Product type", "Product",
    "Net quantity", "Gross sales", "Discounts", "Returns",
    "Net sales", "Shipping", "Taxes", "Total sales",
]

NUMERIC_COLUMNS = [
    "Net quantity", "Gross sales", "Discounts", "Returns",
    "Net sales", "Shipping", "Taxes", "Total sales",
]

# ============================================================
# STYLING
# ============================================================
st.markdown(
    """
    <style>
    :root {
        --navy: #173b65;
        --navy-dark: #102a43;
        --teal: #0d9488;
        --page: #f4f7fb;
        --border: #dce5ef;
        --muted: #52657a;
    }

    .stApp {
        background: var(--page);
        color: var(--navy-dark);
    }

    .block-container {
        max-width: 1550px;
        padding: 1.6rem 2rem 3rem;
    }

    section[data-testid="stSidebar"] {
        background: #ffffff;
        border-right: 1px solid var(--border);
    }

    section[data-testid="stSidebar"] h1,
    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3 {
        color: var(--navy) !important;
    }

    .hero {
        padding: 32px 36px;
        border-radius: 20px;
        background: linear-gradient(
            120deg,
            #102a43,
            #173b65 55%,
            #087e8b
        );
        color: #ffffff;
        margin-bottom: 24px;
        box-shadow: 0 10px 28px rgba(23,59,101,.12);
    }

    .hero h1 {
        color: #ffffff !important;
        font-size: 2.25rem;
        font-weight: 750;
        margin: 0 0 10px 0;
    }

    .hero p {
        color: #e3eef7 !important;
        margin: 0;
        line-height: 1.65;
    }

    h1, h2, h3 {
        color: var(--navy-dark);
    }

    /* Preserve visible KPI labels and values. */
    div[data-testid="stMetric"] {
        background: #ffffff !important;
        border: 1px solid var(--border) !important;
        border-radius: 16px !important;
        padding: 20px 22px !important;
        min-height: 112px;
        box-shadow: 0 4px 14px rgba(15,35,60,.035);
    }

    div[data-testid="stMetric"] * {
        opacity: 1 !important;
        visibility: visible !important;
        text-shadow: none !important;
    }

    div[data-testid="stMetric"] label,
    div[data-testid="stMetric"] label *,
    div[data-testid="stMetricLabel"],
    div[data-testid="stMetricLabel"] *,
    div[data-testid="stMetricLabel"] p,
    div[data-testid="stMetricLabel"] span {
        color: #173b65 !important;
        -webkit-text-fill-color: #173b65 !important;
        font-size: 1rem !important;
        font-weight: 750 !important;
    }

    div[data-testid="stMetricValue"],
    div[data-testid="stMetricValue"] *,
    div[data-testid="stMetricValue"] p {
        color: #102a43 !important;
        -webkit-text-fill-color: #102a43 !important;
        font-size: 1.65rem !important;
        font-weight: 800 !important;
    }

    div[data-testid="stMetricDelta"],
    div[data-testid="stMetricDelta"] * {
        font-weight: 650 !important;
    }

    div[data-testid="stPlotlyChart"] {
        background: #ffffff;
        border: 1px solid var(--border);
        border-radius: 16px;
        padding: 10px;
        box-shadow: 0 4px 14px rgba(15,35,60,.025);
    }

    div[data-testid="stDataFrame"],
    div[data-testid="stTable"] {
        background: #ffffff;
        border: 1px solid var(--border);
        border-radius: 12px;
        overflow: hidden;
    }

    div[data-baseweb="select"] > div,
    div[data-baseweb="input"] > div {
        border-radius: 9px;
    }

    .stButton > button,
    .stDownloadButton > button {
        border-radius: 10px;
        font-weight: 650;
        padding: .55rem 1rem;
    }

    .stButton > button[kind="primary"] {
        background: var(--navy);
        border-color: var(--navy);
        color: #fff;
    }

    .stDownloadButton > button {
        background: #fff;
        border: 1px solid #cbd5e1;
        color: var(--navy);
    }

    .section-caption {
        color: var(--muted);
        font-size: .9rem;
    }

    hr {
        border-color: var(--border);
        margin: 1.5rem 0;
    }

    footer {
        visibility: hidden;
    }

    @media (max-width: 768px) {
        .block-container {
            padding: 1rem;
        }

        .hero {
            padding: 24px 22px;
        }

        div[data-testid="stMetric"] {
            padding: 15px !important;
        }
    }
/* ===== SIDEBAR VISIBILITY FIX ===== */

/* Main sidebar text */
section[data-testid="stSidebar"],
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] span,
section[data-testid="stSidebar"] div {
    color: #173b65;
}

/* Sidebar headings */
section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3 {
    color: #173b65 !important;
}

/* Navigation radio labels */
section[data-testid="stSidebar"] div[role="radiogroup"] label {
    color: #173b65 !important;
    background-color: transparent;
}

/* Selected navigation option */
section[data-testid="stSidebar"] div[role="radiogroup"]
label:has(input:checked) {
    background-color: #e7f3f8 !important;
    border-radius: 8px;
}

/* Filter labels and descriptions */
section[data-testid="stSidebar"] .stMarkdown p,
section[data-testid="stSidebar"] .stWidgetLabel,
section[data-testid="stSidebar"] [data-testid="stWidgetLabel"] {
    color: #52657a !important;
}

/* Date input */
section[data-testid="stSidebar"] input {
    color: #173b65 !important;
}

/* Sidebar buttons */
section[data-testid="stSidebar"] button {
    color: #173b65 !important;
}
    </style>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# DATA LOADING AND CLEANING
# ============================================================
@st.cache_data(show_spinner="Loading and preparing sales data...")
def load_data(file_path, modified_time):
    path = Path(file_path)

    if not path.exists():
        raise FileNotFoundError(f"Dataset not found: {path}")

    df = pd.read_excel(
        path,
        sheet_name=SHEET_NAME,
        engine="openpyxl",
    )

    original_rows = len(df)

    missing_columns = [
        column for column in REQUIRED_COLUMNS
        if column not in df.columns
    ]

    if missing_columns:
        raise ValueError(
            "Missing required columns: " + ", ".join(missing_columns)
        )

    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    invalid_dates = int(df["Date"].isna().sum())

    for column in NUMERIC_COLUMNS:
        df[column] = pd.to_numeric(df[column], errors="coerce")

    missing_financial = int(
        df[["Net sales", "Total sales"]].isna().any(axis=1).sum()
    )

    df = df.dropna(
        subset=["Date", "Net sales", "Total sales"]
    ).copy()

    text_columns = [
        "Order ID",
        "Sale ID",
        "Order",
        "Transaction type",
        "Sale type",
        "Sales channel",
        "Product type",
        "Product",
    ]

    for column in text_columns:
        df[column] = (
            df[column]
            .astype("string")
            .fillna("Unknown")
            .str.strip()
            .replace("", "Unknown")
            .astype(str)
        )

    df = df.sort_values("Date").reset_index(drop=True)

    quality = {
        "original_rows": original_rows,
        "rows_loaded": len(df),
        "rows_excluded": original_rows - len(df),
        "invalid_date_rows": invalid_dates,
        "missing_financial_rows": missing_financial,
        "exact_duplicates": int(df.duplicated().sum()),
        "missing_product_categories": int(
            df["Product type"].eq("Unknown").sum()
        ),
        "missing_products": int(
            df["Product"].eq("Unknown").sum()
        ),
    }

    return df, quality


if not DATA_FILE.exists():
    st.error(
        f"Dataset not found: {DATA_FILE}\n\n"
        "Keep kenyan_ecommerce_sales.xlsx in the same folder as app.py."
    )
    st.stop()

try:
    data, quality = load_data(
        str(DATA_FILE),
        DATA_FILE.stat().st_mtime,
    )
except Exception as error:
    st.error(f"Unable to load the sales dataset: {error}")
    st.stop()

if data.empty:
    st.error("The dataset contains no usable sales records.")
    st.stop()

# ============================================================
# DATE BOUNDARIES
# ============================================================
minimum_date = data["Date"].min().date()
maximum_date = data["Date"].max().date()

# ============================================================
# DASHBOARD HEADER
# ============================================================
st.markdown(
    """
    <div class="hero">
      <h1>📊 Sales Intelligence Dashboard</h1>
      <p>Explore sales performance, product trends, transactions,
      discounts, and returns using the actual sales dataset.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

st.caption(
    f"Dataset period: {minimum_date:%d %b %Y} – "
    f"{maximum_date:%d %b %Y} | {len(data):,} usable rows loaded"
)

# ============================================================
# PREMIUM WORKSPACE NAVIGATION
# ============================================================
st.sidebar.markdown("### WORKSPACE")

page = st.sidebar.radio(
    "Navigate dashboard",
    [
        "Sales Overview",
        "Executive Insights",
        "Product Analysis",
        "Returns & Finance",
        "Data Quality",
        "Transactions",
    ],
    key="workspace_page",
    label_visibility="collapsed",
)
st.markdown(f"**Workspace / {page}**")
st.sidebar.divider()

# ============================================================
# FILTER CALLBACK AND SIDEBAR
# ============================================================
def reset_all_filters():
    st.session_state["date_range"] = (minimum_date, maximum_date)

    for key in (
        "filter_channels",
        "filter_transaction_type",
        "filter_sale_type",
        "filter_product_type",
        "filter_product",
    ):
        st.session_state[key] = []


st.sidebar.title("Dashboard Controls")
st.sidebar.caption("Use these filters to explore the dataset.")
st.sidebar.subheader("Date Range")

selected_dates = st.sidebar.date_input(
    "Select start and end dates",
    value=(minimum_date, maximum_date),
    min_value=minimum_date,
    max_value=maximum_date,
    key="date_range",
)

if isinstance(selected_dates, (tuple, list)):
    if len(selected_dates) == 2:
        start_date, end_date = selected_dates
    elif len(selected_dates) == 1:
        start_date, end_date = selected_dates[0], maximum_date
    else:
        start_date, end_date = minimum_date, maximum_date
else:
    start_date, end_date = selected_dates, maximum_date

if start_date > end_date:
    start_date, end_date = end_date, start_date


def get_options(frame, column):
    return sorted(
        frame[column].dropna().astype(str).unique().tolist()
    )


def select_filter(label, column, key):
    return st.sidebar.multiselect(
        label,
        options=get_options(data, column),
        default=[],
        help="Leave empty to include all options.",
        key=key,
    )


selected_channels = select_filter(
    "Sales Channel",
    "Sales channel",
    "filter_channels",
)

selected_transactions = select_filter(
    "Transaction Type",
    "Transaction type",
    "filter_transaction_type",
)

selected_sale_types = select_filter(
    "Sale Type",
    "Sale type",
    "filter_sale_type",
)

selected_categories = select_filter(
    "Product Category",
    "Product type",
    "filter_product_type",
)

selected_products = select_filter(
    "Product",
    "Product",
    "filter_product",
)

st.sidebar.divider()

st.sidebar.button(
    "Reset All Filters",
    on_click=reset_all_filters,
    use_container_width=True,
)

# ============================================================
# APPLY FILTERS
# ============================================================
filtered_data = data[
    data["Date"].dt.date.between(start_date, end_date)
].copy()

if selected_channels:
    filtered_data = filtered_data[
        filtered_data["Sales channel"].isin(selected_channels)
    ]

if selected_transactions:
    filtered_data = filtered_data[
        filtered_data["Transaction type"].isin(selected_transactions)
    ]

if selected_sale_types:
    filtered_data = filtered_data[
        filtered_data["Sale type"].isin(selected_sale_types)
    ]

if selected_categories:
    filtered_data = filtered_data[
        filtered_data["Product type"].isin(selected_categories)
    ]

if selected_products:
    filtered_data = filtered_data[
        filtered_data["Product"].isin(selected_products)
    ]

# Stop the analytical pages from trying to chart an empty selection.
if filtered_data.empty:
    st.warning(
        "No records match the current filters. "
        "Change your selections or click Reset All Filters."
    )
    st.stop()

# ============================================================
# SHARED METRICS AND PREPARED DATA
# ============================================================
total_net_sales = filtered_data["Net sales"].sum()
total_sales = filtered_data["Total sales"].sum()
total_orders = filtered_data["Order ID"].nunique()
total_quantity = filtered_data["Net quantity"].sum()

average_order_value = (
    total_net_sales / total_orders if total_orders else 0
)

product_data = filtered_data[
    filtered_data["Transaction type"].str.lower().eq("product")
].copy()

# Monthly summary used by sales trends, finance, and exports.
monthly_data = filtered_data.copy()
monthly_data["Month"] = (
    monthly_data["Date"].dt.to_period("M").dt.to_timestamp()
)

monthly_sales = (
    monthly_data.groupby("Month", as_index=False)
    .agg(
        Net_Sales=("Net sales", "sum"),
        Total_Sales=("Total sales", "sum"),
        Orders=("Order ID", "nunique"),
    )
    .sort_values("Month")
)

monthly_sales["Previous_Month_Sales"] = (
    monthly_sales["Net_Sales"].shift(1)
)

previous = monthly_sales["Previous_Month_Sales"]

monthly_sales["MoM_Growth_Percent"] = (
    (
        monthly_sales["Net_Sales"] - previous
    )
    / previous.abs().replace(0, float("nan"))
) * 100

# Shared chart styling.
CHART_TEMPLATE = "plotly_white"

CHART_LAYOUT = {
    "paper_bgcolor": "rgba(0,0,0,0)",
    "plot_bgcolor": "#ffffff",
    "font": {
        "family": "Arial, sans-serif",
        "color": "#173b65",
    },
    "margin": {
        "l": 20,
        "r": 20,
        "t": 55,
        "b": 20,
    },
    "legend": {
        "orientation": "h",
        "yanchor": "bottom",
        "y": 1.02,
        "xanchor": "left",
        "x": 0,
    },
}

# Currency-style formatting without assuming a currency symbol.
def format_amount(value):
    return f"{value:,.2f}"


def apply_chart_layout(fig, title=None):
    fig.update_layout(**CHART_LAYOUT, template=CHART_TEMPLATE)

    if title:
        fig.update_layout(title=title)

    return fig


# ============================================================
# PAGE CONTENT
# ============================================================

# ============================================================
# PAGE 1: SALES OVERVIEW
# ============================================================
if page == "Sales Overview":

    st.title("Sales Performance Overview")
    st.caption(
        "Monitor revenue, order activity, sales trends, and channel performance."
    )

    # KPI cards
    kpi1, kpi2, kpi3, kpi4 = st.columns(4)

    kpi1.metric(
        "Net Sales",
        format_amount(total_net_sales),
        help="Sum of net sales after recorded discounts and returns.",
    )

    kpi2.metric(
        "Total Sales",
        format_amount(total_sales),
        help="Sum of the Total sales column.",
    )

    kpi3.metric(
        "Unique Orders",
        f"{total_orders:,}",
        help="Number of distinct Order IDs in the filtered dataset.",
    )

    kpi4.metric(
        "Net Quantity",
        f"{total_quantity:,.0f}",
        help="Sum of net quantity in the filtered dataset.",
    )

    st.divider()

    # Sales trend
    st.subheader("Sales Trend Over Time")

    if not monthly_sales.empty:
        trend_metric = st.radio(
            "Choose trend metric",
            ["Net Sales", "Total Sales", "Orders"],
            horizontal=True,
            key="overview_trend_metric",
        )

        trend_column = {
            "Net Sales": "Net_Sales",
            "Total Sales": "Total_Sales",
            "Orders": "Orders",
        }[trend_metric]

        trend_fig = px.line(
            monthly_sales,
            x="Month",
            y=trend_column,
            markers=True,
            title=f"Monthly {trend_metric}",
            labels={
                "Month": "Month",
                trend_column: trend_metric,
            },
        )

        trend_fig.update_traces(
            line=dict(width=3),
            marker=dict(size=7),
            hovertemplate=(
                "%{x|%b %Y}<br>"
                + trend_metric
                + ": %{y:,.2f}<extra></extra>"
            ),
        )

        apply_chart_layout(trend_fig)

        st.plotly_chart(
            trend_fig,
            use_container_width=True,
            key="overview_sales_trend",
        )

    st.divider()

    # Channel and transaction analysis
    left, right = st.columns(2)

    with left:
        st.subheader("Sales by Channel")

        channel_data = (
            filtered_data.groupby(
                "Sales channel",
                as_index=False,
            )
            .agg(
                Net_Sales=("Net sales", "sum"),
                Orders=("Order ID", "nunique"),
            )
            .sort_values("Net_Sales", ascending=False)
        )

        if not channel_data.empty:
            channel_fig = px.bar(
                channel_data,
                x="Sales channel",
                y="Net_Sales",
                color="Sales channel",
                title="Net Sales by Sales Channel",
                labels={
                    "Sales channel": "Sales Channel",
                    "Net_Sales": "Net Sales",
                },
                hover_data={"Orders": True},
            )

            apply_chart_layout(channel_fig)

            st.plotly_chart(
                channel_fig,
                use_container_width=True,
                key="overview_channel_chart",
            )

    with right:
        st.subheader("Transaction Breakdown")

        transaction_data = (
            filtered_data.groupby(
                "Transaction type",
                as_index=False,
            )
            .agg(
                Net_Sales=("Net sales", "sum"),
                Records=("Transaction type", "size"),
            )
            .sort_values("Records", ascending=False)
        )

        if not transaction_data.empty:
            transaction_fig = px.pie(
                transaction_data,
                names="Transaction type",
                values="Records",
                hole=0.55,
                title="Share of Transaction Records",
            )

            transaction_fig.update_traces(
                textposition="inside",
                textinfo="percent+label",
            )

            apply_chart_layout(transaction_fig)

            st.plotly_chart(
                transaction_fig,
                use_container_width=True,
                key="overview_transaction_chart",
            )

    st.divider()

    # Sales by sale type
    st.subheader("Sales by Sale Type")

    sale_type_data = (
        filtered_data.groupby(
            "Sale type",
            as_index=False,
        )
        .agg(
            Net_Sales=("Net sales", "sum"),
            Total_Sales=("Total sales", "sum"),
            Orders=("Order ID", "nunique"),
        )
        .sort_values("Net_Sales", ascending=False)
    )

    if not sale_type_data.empty:
        sale_type_fig = px.bar(
            sale_type_data,
            x="Sale type",
            y="Net_Sales",
            color="Sale type",
            title="Net Sales by Sale Type",
            labels={
                "Sale type": "Sale Type",
                "Net_Sales": "Net Sales",
            },
            hover_data={
                "Total_Sales": ":,.2f",
                "Orders": True,
            },
        )

        apply_chart_layout(sale_type_fig)

        st.plotly_chart(
            sale_type_fig,
            use_container_width=True,
            key="overview_sale_type_chart",
        )

    st.subheader("Recent Sales Records")

    recent_columns = [
        "Date",
        "Order ID",
        "Transaction type",
        "Sales channel",
        "Product",
        "Net quantity",
        "Net sales",
        "Total sales",
    ]

    recent_data = (
        filtered_data[recent_columns]
        .sort_values("Date", ascending=False)
        .head(10)
        .copy()
    )

    recent_data["Date"] = recent_data["Date"].dt.strftime("%d %b %Y")

    st.dataframe(
        recent_data,
        use_container_width=True,
        hide_index=True,
    )


# ============================================================
# PAGE 2: ADVANCED PRODUCT INTELLIGENCE
# ============================================================
elif page == "Product Analysis":

    st.title("Product Intelligence")
    st.caption(
        "Explore product rankings, revenue contribution, category "
        "performance, monthly trends, and records requiring investigation."
    )

    # --------------------------------------------------------
    # PREPARE PRODUCT TRANSACTIONS
    # --------------------------------------------------------
    product_data = filtered_data[
        filtered_data["Transaction type"].str.lower().eq("product")
    ].copy()

    if product_data.empty:
        st.warning(
            "No product transactions match the current filters. "
            "Change the Transaction Type filter in the sidebar."
        )

    else:
        # Product-level aggregation
        product_ranking = (
            product_data.groupby(
                ["Product", "Product type"],
                as_index=False,
                dropna=False,
            )
            .agg(
                Net_Sales=("Net sales", "sum"),
                Gross_Sales=("Gross sales", "sum"),
                Discounts=("Discounts", "sum"),
                Returns=("Returns", "sum"),
                Total_Sales=("Total sales", "sum"),
                Net_Quantity=("Net quantity", "sum"),
                Orders=("Order ID", "nunique"),
                Transaction_Count=("Order ID", "size"),
            )
            .sort_values("Net_Sales", ascending=False)
            .reset_index(drop=True)
        )

        # Category-level aggregation
        category_data = (
            product_data.groupby(
                "Product type",
                as_index=False,
                dropna=False,
            )
            .agg(
                Net_Sales=("Net sales", "sum"),
                Gross_Sales=("Gross sales", "sum"),
                Net_Quantity=("Net quantity", "sum"),
                Orders=("Order ID", "nunique"),
                Products=("Product", "nunique"),
            )
            .sort_values("Net_Sales", ascending=False)
        )

        total_product_sales = product_ranking["Net_Sales"].sum()
        total_product_quantity = product_ranking["Net_Quantity"].sum()
        total_unique_products = product_ranking["Product"].nunique()
        total_categories = product_ranking["Product type"].nunique()

        # ----------------------------------------------------
        # 1. PRODUCT KPI CARDS
        # ----------------------------------------------------
        st.subheader("Product Performance at a Glance")

        p1, p2, p3, p4 = st.columns(4)

        p1.metric(
            "Product Net Sales",
            format_amount(total_product_sales),
        )

        p2.metric(
            "Net Quantity",
            f"{total_product_quantity:,.0f}",
        )

        p3.metric(
            "Unique Products",
            f"{total_unique_products:,}",
        )

        p4.metric(
            "Product Categories",
            f"{total_categories:,}",
        )

        st.divider()

        # ----------------------------------------------------
        # 2. PRODUCT RANKING CONTROLS
        # ----------------------------------------------------
        st.subheader("Product Rankings")

        ranking_options = st.columns(2)

        with ranking_options[0]:
            ranking_metric = st.selectbox(
                "Rank products by",
                [
                    "Net Sales",
                    "Gross Sales",
                    "Total Sales",
                    "Net Quantity",
                    "Unique Orders",
                ],
                key="advanced_product_ranking_metric",
            )

        with ranking_options[1]:
            ranking_count = st.selectbox(
                "Number of products",
                [5, 10, 15, 20],
                index=1,
                key="advanced_product_ranking_count",
            )

        ranking_column = {
            "Net Sales": "Net_Sales",
            "Gross Sales": "Gross_Sales",
            "Total Sales": "Total_Sales",
            "Net Quantity": "Net_Quantity",
            "Unique Orders": "Orders",
        }[ranking_metric]

        top_products = (
            product_ranking
            .sort_values(ranking_column, ascending=False)
            .head(ranking_count)
            .copy()
        )

        bottom_products = (
            product_ranking
            .sort_values(ranking_column, ascending=True)
            .head(ranking_count)
            .copy()
        )

        top_left, top_right = st.columns(2)

        with top_left:
            st.markdown(f"#### Top {ranking_count} Products")

            top_chart = px.bar(
                top_products.sort_values(ranking_column),
                x=ranking_column,
                y="Product",
                color="Product type",
                orientation="h",
                title=f"Top Products by {ranking_metric}",
                labels={
                    ranking_column: ranking_metric,
                    "Product": "Product",
                    "Product type": "Category",
                },
                hover_data={
                    "Net_Sales": ":,.2f",
                    "Net_Quantity": ":,.0f",
                    "Orders": True,
                },
            )

            apply_chart_layout(top_chart)

            st.plotly_chart(
                top_chart,
                use_container_width=True,
                key="advanced_product_top_chart",
            )

        with top_right:
            st.markdown(f"#### Bottom {ranking_count} Products")

            bottom_chart = px.bar(
                bottom_products.sort_values(ranking_column),
                x=ranking_column,
                y="Product",
                color="Product type",
                orientation="h",
                title=f"Bottom Products by {ranking_metric}",
                labels={
                    ranking_column: ranking_metric,
                    "Product": "Product",
                    "Product type": "Category",
                },
                hover_data={
                    "Net_Sales": ":,.2f",
                    "Net_Quantity": ":,.0f",
                    "Orders": True,
                },
            )

            apply_chart_layout(bottom_chart)

            st.plotly_chart(
                bottom_chart,
                use_container_width=True,
                key="advanced_product_bottom_chart",
            )

        st.caption(
            "The bottom ranking identifies the lowest recorded values "
            "for the selected metric. Low sales alone do not establish "
            "that a product is unprofitable."
        )

        st.divider()

        # ----------------------------------------------------
        # 3. PRODUCT SALES CONTRIBUTION
        # ----------------------------------------------------
        st.subheader("Product Contribution to Net Sales")

        contribution_data = product_ranking.copy()

        if total_product_sales != 0:
            contribution_data["Sales_Contribution_Percent"] = (
                contribution_data["Net_Sales"]
                / total_product_sales
            ) * 100
        else:
            contribution_data["Sales_Contribution_Percent"] = 0.0

        contribution_data = contribution_data.sort_values(
            "Sales_Contribution_Percent",
            ascending=False,
        )

        contribution_left, contribution_right = st.columns(2)

        with contribution_left:
            contribution_count = st.slider(
                "Products to display in contribution chart",
                min_value=5,
                max_value=min(
                    30,
                    max(5, len(contribution_data)),
                ),
                value=min(10, max(5, len(contribution_data))),
                key="product_contribution_count",
            )

            contribution_top = contribution_data.head(
                contribution_count
            )

            contribution_chart = px.bar(
                contribution_top.sort_values(
                    "Sales_Contribution_Percent"
                ),
                x="Sales_Contribution_Percent",
                y="Product",
                orientation="h",
                color="Product type",
                title="Product Share of Net Sales",
                labels={
                    "Sales_Contribution_Percent": "Contribution (%)",
                    "Product": "Product",
                    "Product type": "Category",
                },
            )

            apply_chart_layout(contribution_chart)

            st.plotly_chart(
                contribution_chart,
                use_container_width=True,
                key="product_contribution_chart",
            )

        with contribution_right:
            category_chart = px.bar(
                category_data.sort_values("Net_Sales"),
                x="Net_Sales",
                y="Product type",
                orientation="h",
                color="Product type",
                title="Category Net Sales",
                labels={
                    "Net_Sales": "Net Sales",
                    "Product type": "Category",
                },
                hover_data={
                    "Net_Quantity": ":,.0f",
                    "Orders": True,
                    "Products": True,
                },
            )

            apply_chart_layout(category_chart)

            st.plotly_chart(
                category_chart,
                use_container_width=True,
                key="advanced_category_chart",
            )

        st.divider()

        # ----------------------------------------------------
        # 4. PRODUCT DRILL-DOWN
        # ----------------------------------------------------
        st.subheader("Product Drill-Down")

        available_products = sorted(
            product_data["Product"].dropna().unique().tolist()
        )

        selected_product = st.selectbox(
            "Select a product to investigate",
            available_products,
            key="advanced_selected_product",
        )

        selected_product_data = product_data[
            product_data["Product"].eq(selected_product)
        ].copy()

        if not selected_product_data.empty:

            selected_product_categories = (
                selected_product_data["Product type"]
                .dropna()
                .unique()
                .tolist()
            )

            selected_product_category = (
                selected_product_categories[0]
                if selected_product_categories
                else "Unknown"
            )

            selected_net_sales = selected_product_data["Net sales"].sum()
            selected_quantity = selected_product_data["Net quantity"].sum()
            selected_orders = selected_product_data["Order ID"].nunique()

            d1, d2, d3, d4 = st.columns(4)

            d1.metric(
                "Selected Product",
                selected_product,
            )

            d2.metric(
                "Net Sales",
                format_amount(selected_net_sales),
            )

            d3.metric(
                "Net Quantity",
                f"{selected_quantity:,.0f}",
            )

            d4.metric(
                "Unique Orders",
                f"{selected_orders:,}",
            )

            selected_product_data["Month"] = (
                selected_product_data["Date"]
                .dt.to_period("M")
                .dt.to_timestamp()
            )

            product_monthly = (
                selected_product_data.groupby(
                    "Month",
                    as_index=False,
                )
                .agg(
                    Net_Sales=("Net sales", "sum"),
                    Net_Quantity=("Net quantity", "sum"),
                    Orders=("Order ID", "nunique"),
                )
                .sort_values("Month")
            )

            monthly_product_chart = px.line(
                product_monthly,
                x="Month",
                y="Net_Sales",
                markers=True,
                title=f"Monthly Net Sales — {selected_product}",
                labels={
                    "Month": "Month",
                    "Net_Sales": "Net Sales",
                },
            )

            monthly_product_chart.update_traces(
                line=dict(width=3),
                marker=dict(size=7),
            )

            apply_chart_layout(monthly_product_chart)

            st.plotly_chart(
                monthly_product_chart,
                use_container_width=True,
                key="selected_product_monthly_chart",
            )

            st.markdown("#### Selected Product Transactions")

            product_transaction_columns = [
                "Date",
                "Order ID",
                "Sale ID",
                "Sales channel",
                "Net quantity",
                "Net sales",
                "Returns",
                "Total sales",
            ]

            product_transaction_columns = [
                column
                for column in product_transaction_columns
                if column in selected_product_data.columns
            ]

            st.dataframe(
                selected_product_data[
                    product_transaction_columns
                ].sort_values("Date", ascending=False),
                use_container_width=True,
                hide_index=True,
            )

            st.download_button(
                "Download Selected Product Transactions",
                data=selected_product_data.to_csv(
                    index=False,
                    date_format="%Y-%m-%d",
                ).encode("utf-8"),
                file_name="selected_product_transactions.csv",
                mime="text/csv",
                key="download_selected_product_transactions",
            )

        st.divider()

        # ----------------------------------------------------
        # 5. ANOMALY DETECTION
        # ----------------------------------------------------
        st.subheader("Records Requiring Investigation")

        st.caption(
            "These checks flag unusual recorded values. They do not "
            "automatically indicate errors or fraud."
        )

        negative_sales = product_data[
            product_data["Net sales"] < 0
        ].copy()

        negative_quantity = product_data[
            product_data["Net quantity"] < 0
        ].copy()

        negative_returns = product_data[
            product_data["Returns"] < 0
        ].copy()

        a1, a2, a3 = st.columns(3)

        a1.metric(
            "Negative Net Sales Records",
            f"{len(negative_sales):,}",
        )

        a2.metric(
            "Negative Quantity Records",
            f"{len(negative_quantity):,}",
        )

        a3.metric(
            "Negative Returns Records",
            f"{len(negative_returns):,}",
        )

        with st.expander(
            "Inspect records with negative net sales",
            expanded=False,
        ):
            if negative_sales.empty:
                st.success("No negative net sales records were found.")
            else:
                st.dataframe(
                    negative_sales,
                    use_container_width=True,
                    hide_index=True,
                )

                st.download_button(
                    "Download Negative Net Sales Records",
                    data=negative_sales.to_csv(
                        index=False,
                        date_format="%Y-%m-%d",
                    ).encode("utf-8"),
                    file_name="negative_net_sales_records.csv",
                    mime="text/csv",
                    key="download_negative_net_sales",
                )

        with st.expander(
            "Inspect records with negative quantities",
            expanded=False,
        ):
            if negative_quantity.empty:
                st.success("No negative quantity records were found.")
            else:
                st.dataframe(
                    negative_quantity,
                    use_container_width=True,
                    hide_index=True,
                )

                st.download_button(
                    "Download Negative Quantity Records",
                    data=negative_quantity.to_csv(
                        index=False,
                        date_format="%Y-%m-%d",
                    ).encode("utf-8"),
                    file_name="negative_quantity_records.csv",
                    mime="text/csv",
                    key="download_negative_quantity",
                )

        with st.expander(
            "Inspect records with negative returns",
            expanded=False,
        ):
            if negative_returns.empty:
                st.success("No negative returns records were found.")
            else:
                st.dataframe(
                    negative_returns,
                    use_container_width=True,
                    hide_index=True,
                )

                st.download_button(
                    "Download Negative Returns Records",
                    data=negative_returns.to_csv(
                        index=False,
                        date_format="%Y-%m-%d",
                    ).encode("utf-8"),
                    file_name="negative_returns_records.csv",
                    mime="text/csv",
                    key="download_negative_returns",
                )

        st.divider()

        # ----------------------------------------------------
        # 6. EXPORT PRODUCT INTELLIGENCE
        # ----------------------------------------------------
        st.subheader("Export Product Intelligence")

        export_left, export_right = st.columns(2)

        with export_left:
            st.download_button(
                "Download Full Product Rankings",
                data=product_ranking.to_csv(
                    index=False
                ).encode("utf-8"),
                file_name="product_rankings.csv",
                mime="text/csv",
                key="download_advanced_product_rankings",
                use_container_width=True,
            )

        with export_right:
            st.download_button(
                "Download Category Analysis",
                data=category_data.to_csv(
                    index=False
                ).encode("utf-8"),
                file_name="category_analysis.csv",
                mime="text/csv",
                key="download_advanced_category_analysis",
                use_container_width=True,
            )

# ============================================================
# PAGE 3: RETURNS & FINANCE
# ============================================================
elif page == "Returns & Finance":

    st.title("Returns & Financial Analysis")
    st.caption(
        "Review sales totals, discounts, returns, shipping, taxes, "
        "and monthly financial trends."
    )

    # Financial KPIs
    f1, f2, f3, f4 = st.columns(4)

    f1.metric(
        "Gross Sales",
        format_amount(filtered_data["Gross sales"].sum()),
    )

    f2.metric(
        "Discounts",
        format_amount(filtered_data["Discounts"].sum()),
    )

    f3.metric(
        "Returns",
        format_amount(filtered_data["Returns"].sum()),
    )

    f4.metric(
        "Shipping",
        format_amount(filtered_data["Shipping"].sum()),
    )

    st.divider()

    # Financial totals
    st.subheader("Financial Summary")

    financial_summary = pd.DataFrame(
        {
            "Metric": [
                "Gross Sales",
                "Discounts",
                "Returns",
                "Net Sales",
                "Shipping",
                "Taxes",
                "Total Sales",
            ],
            "Amount": [
                filtered_data["Gross sales"].sum(),
                filtered_data["Discounts"].sum(),
                filtered_data["Returns"].sum(),
                filtered_data["Net sales"].sum(),
                filtered_data["Shipping"].sum(),
                filtered_data["Taxes"].sum(),
                filtered_data["Total sales"].sum(),
            ],
        }
    )

    finance_left, finance_right = st.columns([1, 1.5])

    with finance_left:
        st.dataframe(
            financial_summary,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Amount": st.column_config.NumberColumn(
                    "Amount",
                    format="%,.2f",
                ),
            },
        )

    with finance_right:
        finance_fig = px.bar(
            financial_summary,
            x="Metric",
            y="Amount",
            color="Metric",
            title="Financial Components",
            labels={
                "Metric": "Financial Metric",
                "Amount": "Amount",
            },
        )

        apply_chart_layout(finance_fig)

        st.plotly_chart(
            finance_fig,
            use_container_width=True,
            key="finance_components_chart",
        )

    st.divider()

    # Monthly financial trend
    st.subheader("Monthly Financial Performance")

    if not monthly_sales.empty:
        finance_metric = st.selectbox(
            "Choose financial trend",
            ["Net Sales", "Total Sales"],
            key="finance_trend_metric",
        )

        finance_column = {
            "Net Sales": "Net_Sales",
            "Total Sales": "Total_Sales",
        }[finance_metric]

        finance_trend_fig = px.line(
            monthly_sales,
            x="Month",
            y=finance_column,
            markers=True,
            title=f"Monthly {finance_metric}",
            labels={
                "Month": "Month",
                finance_column: finance_metric,
            },
        )

        finance_trend_fig.update_traces(line=dict(width=3))

        apply_chart_layout(finance_trend_fig)

        st.plotly_chart(
            finance_trend_fig,
            use_container_width=True,
            key="finance_monthly_trend",
        )

        # Monthly growth analysis
        st.subheader("Month-over-Month Sales Growth")

        growth_data = monthly_sales.dropna(
            subset=["MoM_Growth_Percent"]
        ).copy()

        if not growth_data.empty:
            growth_fig = px.bar(
                growth_data,
                x="Month",
                y="MoM_Growth_Percent",
                title="Monthly Net Sales Growth",
                labels={
                    "Month": "Month",
                    "MoM_Growth_Percent": "Growth (%)",
                },
            )

            growth_fig.add_hline(
                y=0,
                line_dash="dash",
                line_color="#64748b",
            )

            apply_chart_layout(growth_fig)

            st.plotly_chart(
                growth_fig,
                use_container_width=True,
                key="finance_growth_chart",
            )

        st.caption(
            "Month-over-month growth compares the current month's net "
            "sales with the previous month's net sales. The first month "
            "has no previous month for comparison."
        )

    st.divider()

    # Sales channel financial comparison
    st.subheader("Financial Performance by Channel")

    channel_finance = (
        filtered_data.groupby(
            "Sales channel",
            as_index=False,
        )
        .agg(
            Gross_Sales=("Gross sales", "sum"),
            Discounts=("Discounts", "sum"),
            Returns=("Returns", "sum"),
            Net_Sales=("Net sales", "sum"),
            Total_Sales=("Total sales", "sum"),
        )
        .sort_values("Net_Sales", ascending=False)
    )

    st.dataframe(
        channel_finance,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Gross_Sales": st.column_config.NumberColumn(
                "Gross Sales",
                format="%,.2f",
            ),
            "Discounts": st.column_config.NumberColumn(
                "Discounts",
                format="%,.2f",
            ),
            "Returns": st.column_config.NumberColumn(
                "Returns",
                format="%,.2f",
            ),
            "Net_Sales": st.column_config.NumberColumn(
                "Net Sales",
                format="%,.2f",
            ),
            "Total_Sales": st.column_config.NumberColumn(
                "Total Sales",
                format="%,.2f",
            ),
        },
    )

    st.download_button(
        "Download Financial Summary CSV",
        data=financial_summary.to_csv(index=False).encode("utf-8"),
        file_name="financial_summary.csv",
        mime="text/csv",
        key="download_financial_summary",
    )

# ============================================================
# PAGE 4: DATA QUALITY
# ============================================================
elif page == "Data Quality":

    st.title("Data Quality & Integrity")
    st.caption(
        "Inspect dataset completeness, missing values, duplicate records, "
        "and the effects of data cleaning."
    )

    # Dataset-level metrics
    q1, q2, q3, q4 = st.columns(4)

    q1.metric(
        "Original Rows",
        f"{quality['original_rows']:,}",
    )

    q2.metric(
        "Usable Rows",
        f"{quality['rows_loaded']:,}",
    )

    q3.metric(
        "Excluded Rows",
        f"{quality['rows_excluded']:,}",
    )

    q4.metric(
        "Exact Duplicates",
        f"{quality['exact_duplicates']:,}",
    )

    st.divider()

    # Cleaning summary
    st.subheader("Data Preparation Summary")

    preparation_data = pd.DataFrame(
        {
            "Quality Check": [
                "Rows in original dataset",
                "Rows retained after cleaning",
                "Rows excluded during cleaning",
                "Rows with invalid dates",
                "Rows with missing financial values",
                "Exact duplicate rows in cleaned data",
                "Rows with unknown product categories",
                "Rows with unknown product names",
            ],
            "Result": [
                quality["original_rows"],
                quality["rows_loaded"],
                quality["rows_excluded"],
                quality["invalid_date_rows"],
                quality["missing_financial_rows"],
                quality["exact_duplicates"],
                quality["missing_product_categories"],
                quality["missing_products"],
            ],
        }
    )

    st.dataframe(
        preparation_data,
        use_container_width=True,
        hide_index=True,
    )

    st.info(
        "The dashboard excludes rows with invalid dates or missing "
        "Net sales or Total sales values. Other missing numeric values "
        "are retained and may be ignored by pandas when calculating sums."
    )

    st.divider()

    # Missing values in the original columns represented by the
    # cleaned data are no longer fully recoverable for text fields,
    # so explain the scope of this diagnostic.
    st.subheader("Current Filtered Dataset Profile")

    profile = pd.DataFrame(
        {
            "Column": filtered_data.columns,
            "Data Type": [
                str(filtered_data[column].dtype)
                for column in filtered_data.columns
            ],
            "Missing Values": [
                int(filtered_data[column].isna().sum())
                for column in filtered_data.columns
            ],
            "Unique Values": [
                int(filtered_data[column].nunique(dropna=True))
                for column in filtered_data.columns
            ],
        }
    )

    st.dataframe(
        profile,
        use_container_width=True,
        hide_index=True,
    )

    st.caption(
        "This profile describes the cleaned dataset under the currently "
        "selected filters. Missing text values were normalized to "
        "'Unknown' during preparation."
    )

    st.divider()

    st.subheader("Download Data Quality Report")

    st.download_button(
        "Download Quality Summary CSV",
        data=preparation_data.to_csv(index=False).encode("utf-8"),
        file_name="data_quality_summary.csv",
        mime="text/csv",
        key="download_quality_summary",
    )

    st.download_button(
        "Download Filtered Data Profile CSV",
        data=profile.to_csv(index=False).encode("utf-8"),
        file_name="filtered_data_profile.csv",
        mime="text/csv",
        key="download_filtered_profile",
    )


# ============================================================
# PAGE 5: TRANSACTIONS
# ============================================================
elif page == "Transactions":

    st.title("Transaction Explorer")
    st.caption(
        "Search, inspect, filter, and export transaction records "
        "from the current dataset selection."
    )

    # Transaction KPIs
    t1, t2, t3, t4 = st.columns(4)

    t1.metric(
        "Matching Records",
        f"{len(filtered_data):,}",
    )

    t2.metric(
        "Unique Orders",
        f"{filtered_data['Order ID'].nunique():,}",
    )

    t3.metric(
        "Unique Products",
        f"{filtered_data['Product'].nunique():,}",
    )

    t4.metric(
        "Net Sales",
        format_amount(filtered_data["Net sales"].sum()),
    )

    st.divider()

    # Search across relevant fields
    st.subheader("Search Transactions")

    search_term = st.text_input(
        "Search by order ID, sale ID, product, channel, or transaction type",
        placeholder="Enter an order ID, product name, or keyword...",
        key="transaction_search",
    )

    transaction_view = filtered_data.copy()

    if search_term.strip():
        searchable_columns = [
            "Order ID",
            "Sale ID",
            "Product",
            "Product type",
            "Sales channel",
            "Transaction type",
            "Sale type",
        ]

        search_mask = pd.Series(
            False,
            index=transaction_view.index,
        )

        for column in searchable_columns:
            search_mask |= (
                transaction_view[column]
                .astype(str)
                .str.contains(
                    search_term.strip(),
                    case=False,
                    na=False,
                    regex=False,
                )
            )

        transaction_view = transaction_view[search_mask].copy()

    st.write(
        f"**{len(transaction_view):,}** records match your search "
        f"within the selected filters."
    )

    # Select columns for the table
    preferred_columns = [
        "Date",
        "Order ID",
        "Sale ID",
        "Transaction type",
        "Sale type",
        "Sales channel",
        "Product type",
        "Product",
        "Net quantity",
        "Gross sales",
        "Discounts",
        "Returns",
        "Net sales",
        "Shipping",
        "Taxes",
        "Total sales",
    ]

    available_columns = [
        column
        for column in preferred_columns
        if column in transaction_view.columns
    ]

    display_transactions = transaction_view[available_columns].copy()

    # Display dates in a readable format.
    if "Date" in display_transactions.columns:
        display_transactions["Date"] = (
            display_transactions["Date"].dt.strftime("%Y-%m-%d")
        )

    # Table pagination
    page_size = st.selectbox(
        "Rows per page",
        [25, 50, 100, 250],
        index=1,
        key="transaction_page_size",
    )

    total_pages = max(
        1,
        (len(display_transactions) + page_size - 1) // page_size,
    )

    current_page = st.number_input(
        "Page",
        min_value=1,
        max_value=total_pages,
        value=1,
        step=1,
        key="transaction_page_number",
    )

    start_index = (current_page - 1) * page_size
    end_index = start_index + page_size

    page_transactions = display_transactions.iloc[
        start_index:end_index
    ]

    st.dataframe(
        page_transactions,
        use_container_width=True,
        hide_index=True,
        height=500,
        column_config={
            "Net quantity": st.column_config.NumberColumn(
                "Net Quantity",
                format="%,.0f",
            ),
            "Gross sales": st.column_config.NumberColumn(
                "Gross Sales",
                format="%,.2f",
            ),
            "Discounts": st.column_config.NumberColumn(
                "Discounts",
                format="%,.2f",
            ),
            "Returns": st.column_config.NumberColumn(
                "Returns",
                format="%,.2f",
            ),
            "Net sales": st.column_config.NumberColumn(
                "Net Sales",
                format="%,.2f",
            ),
            "Shipping": st.column_config.NumberColumn(
                "Shipping",
                format="%,.2f",
            ),
            "Taxes": st.column_config.NumberColumn(
                "Taxes",
                format="%,.2f",
            ),
            "Total sales": st.column_config.NumberColumn(
                "Total Sales",
                format="%,.2f",
            ),
        },
    )

    st.caption(
        f"Showing records {start_index + 1:,}–"
        f"{min(end_index, len(display_transactions)):,} "
        f"of {len(display_transactions):,} matching records."
        if len(display_transactions)
        else "No matching records found."
    )

    st.divider()

    # CSV exports
    st.subheader("Export Transactions")

    export_left, export_right = st.columns(2)

    with export_left:
        st.download_button(
            "Download Matching Transactions CSV",
            data=transaction_view.to_csv(
                index=False,
                date_format="%Y-%m-%d",
            ).encode("utf-8"),
            file_name="matching_transactions.csv",
            mime="text/csv",
            key="download_matching_transactions",
            use_container_width=True,
        )

    with export_right:
        st.download_button(
            "Download Monthly Sales Summary CSV",
            data=monthly_sales.to_csv(
                index=False,
                date_format="%Y-%m-%d",
            ).encode("utf-8"),
            file_name="monthly_sales_summary.csv",
            mime="text/csv",
            key="download_monthly_sales",
            use_container_width=True,
        )

    st.caption(
        "Transaction exports respect the global date, channel, "
        "transaction type, sale type, category, and product filters. "
        "The matching-transactions export also respects the search term."
    )
# ============================================================
# EXECUTIVE INSIGHTS
# ============================================================
elif page == "Executive Insights":

    st.title("Executive Insights")
    st.caption(
        "A decision-focused summary of sales performance, "
        "product contribution, channel performance, and financial trends."
    )

    # --------------------------------------------------------
    # 1. EXECUTIVE KPI SUMMARY
    # --------------------------------------------------------
    st.subheader("Performance at a Glance")

    exec_orders = filtered_data["Order ID"].nunique()
    exec_net_sales = filtered_data["Net sales"].sum()
    exec_total_sales = filtered_data["Total sales"].sum()

    exec_aov = (
        exec_net_sales / exec_orders
        if exec_orders > 0
        else 0
    )

    k1, k2, k3, k4 = st.columns(4)

    k1.metric(
        "Net Sales",
        format_amount(exec_net_sales),
    )

    k2.metric(
        "Total Sales",
        format_amount(exec_total_sales),
    )

    k3.metric(
        "Unique Orders",
        f"{exec_orders:,}",
    )

    k4.metric(
        "Average Net Sales per Order",
        format_amount(exec_aov),
    )

    st.divider()

    # --------------------------------------------------------
    # 2. AUTOMATIC SALES INSIGHTS
    # --------------------------------------------------------
    st.subheader("Key Business Findings")

    insights = []

    # Monthly sales trend
    if len(monthly_sales) >= 2:

        latest_month = monthly_sales.iloc[-1]
        previous_month = monthly_sales.iloc[-2]

        latest_sales = latest_month["Net_Sales"]
        previous_sales = previous_month["Net_Sales"]

        latest_label = latest_month["Month"].strftime("%B %Y")
        previous_label = previous_month["Month"].strftime("%B %Y")

        if previous_sales > 0:

            growth = (
                (latest_sales - previous_sales)
                / abs(previous_sales)
            ) * 100

            if growth > 0:
                insights.append(
                    (
                        "positive",
                        "Monthly sales increased",
                        f"Net sales for {latest_label} were "
                        f"{growth:.1f}% higher than {previous_label}."
                    )
                )

            elif growth < 0:
                insights.append(
                    (
                        "warning",
                        "Monthly sales declined",
                        f"Net sales for {latest_label} were "
                        f"{abs(growth):.1f}% lower than {previous_label}. "
                        "Investigate changes in order volume, product mix, "
                        "and sales channels."
                    )
                )

            else:
                insights.append(
                    (
                        "neutral",
                        "Monthly sales were unchanged",
                        f"Net sales for {latest_label} matched "
                        f"{previous_label}."
                    )
                )

        else:
            insights.append(
                (
                    "neutral",
                    "Monthly comparison needs review",
                    f"Net sales for {previous_label} were zero or negative, "
                    "so a conventional percentage growth rate is not useful."
                )
            )

    elif len(monthly_sales) == 1:
        insights.append(
            (
                "neutral",
                "Limited monthly history",
                "Only one month is available under the current filters. "
                "Expand the date range to compare monthly performance."
            )
        )

    # Leading sales channel
    channel_insights = (
        filtered_data.groupby("Sales channel")["Net sales"]
        .sum()
        .sort_values(ascending=False)
    )

    if not channel_insights.empty:

        leading_channel = channel_insights.index[0]
        leading_channel_sales = channel_insights.iloc[0]

        if exec_net_sales != 0:
            channel_share = (
                leading_channel_sales / exec_net_sales
            ) * 100

            insights.append(
                (
                    "positive",
                    "Leading sales channel",
                    f"{leading_channel} generated "
                    f"{format_amount(leading_channel_sales)} in net sales, "
                    f"equivalent to {channel_share:.1f}% of total net sales."
                )
            )
        else:
            insights.append(
                (
                    "neutral",
                    "Leading sales channel",
                    f"{leading_channel} had the highest net sales "
                    "among the selected channels."
                )
            )

    # Leading product category
    exec_products = filtered_data[
        filtered_data["Transaction type"].str.lower().eq("product")
    ].copy()

    if not exec_products.empty:

        category_insights = (
            exec_products.groupby("Product type")["Net sales"]
            .sum()
            .sort_values(ascending=False)
        )

        if not category_insights.empty:

            leading_category = category_insights.index[0]
            leading_category_sales = category_insights.iloc[0]

            insights.append(
                (
                    "positive",
                    "Leading product category",
                    f"{leading_category} generated "
                    f"{format_amount(leading_category_sales)} in net sales "
                    "within the selected product transactions."
                )
            )

        product_totals = (
            exec_products.groupby("Product")["Net sales"]
            .sum()
            .sort_values(ascending=False)
        )

        if not product_totals.empty:

            best_product = product_totals.index[0]
            best_product_sales = product_totals.iloc[0]

            insights.append(
                (
                    "neutral",
                    "Top-performing product",
                    f"{best_product} recorded the highest product-level "
                    f"net sales: {format_amount(best_product_sales)}."
                )
            )

    # Display findings
    for kind, title, message in insights:

        if kind == "positive":
            st.success(f"**{title}**\n\n{message}")

        elif kind == "warning":
            st.warning(f"**{title}**\n\n{message}")

        else:
            st.info(f"**{title}**\n\n{message}")

    st.divider()

    # --------------------------------------------------------
    # 3. MONTHLY PERFORMANCE
    # --------------------------------------------------------
    st.subheader("Monthly Sales Performance")

    if not monthly_sales.empty:

        executive_trend = px.line(
            monthly_sales,
            x="Month",
            y="Net_Sales",
            markers=True,
            title="Net Sales by Month",
            labels={
                "Month": "Month",
                "Net_Sales": "Net Sales",
            },
        )

        executive_trend.update_traces(
            line=dict(width=3),
            marker=dict(size=7),
        )

        apply_chart_layout(executive_trend)

        st.plotly_chart(
            executive_trend,
            use_container_width=True,
            key="executive_monthly_trend",
        )

        # Monthly performance table
        executive_monthly_table = monthly_sales[
            [
                "Month",
                "Net_Sales",
                "Total_Sales",
                "Orders",
                "MoM_Growth_Percent",
            ]
        ].copy()

        executive_monthly_table["Month"] = (
            executive_monthly_table["Month"].dt.strftime("%b %Y")
        )

        st.dataframe(
            executive_monthly_table,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Month": "Month",
                "Net_Sales": st.column_config.NumberColumn(
                    "Net Sales",
                    format="%,.2f",
                ),
                "Total_Sales": st.column_config.NumberColumn(
                    "Total Sales",
                    format="%,.2f",
                ),
                "Orders": st.column_config.NumberColumn(
                    "Unique Orders",
                    format="%,.0f",
                ),
                "MoM_Growth_Percent": st.column_config.NumberColumn(
                    "Month-over-Month Growth",
                    format="%.2f%%",
                ),
            },
        )

    st.divider()

    # --------------------------------------------------------
    # 4. CHANNEL PERFORMANCE
    # --------------------------------------------------------
    st.subheader("Sales Channel Comparison")

    executive_channels = (
        filtered_data.groupby("Sales channel", as_index=False)
        .agg(
            Net_Sales=("Net sales", "sum"),
            Total_Sales=("Total sales", "sum"),
            Orders=("Order ID", "nunique"),
        )
        .sort_values("Net_Sales", ascending=False)
    )

    if not executive_channels.empty:

        channel_chart = px.bar(
            executive_channels,
            x="Sales channel",
            y="Net_Sales",
            color="Sales channel",
            title="Net Sales Contribution by Channel",
            labels={
                "Sales channel": "Sales Channel",
                "Net_Sales": "Net Sales",
            },
            hover_data={
                "Total_Sales": ":,.2f",
                "Orders": True,
            },
        )

        apply_chart_layout(channel_chart)

        st.plotly_chart(
            channel_chart,
            use_container_width=True,
            key="executive_channel_comparison",
        )

        st.dataframe(
            executive_channels,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Net_Sales": st.column_config.NumberColumn(
                    "Net Sales",
                    format="%,.2f",
                ),
                "Total_Sales": st.column_config.NumberColumn(
                    "Total Sales",
                    format="%,.2f",
                ),
                "Orders": st.column_config.NumberColumn(
                    "Unique Orders",
                    format="%,.0f",
                ),
            },
        )

    st.divider()

    # --------------------------------------------------------
    # 5. FINANCIAL WATCHLIST
    # --------------------------------------------------------
    st.subheader("Financial Watchlist")

    gross_sales = filtered_data["Gross sales"].sum()
    discounts_total = filtered_data["Discounts"].sum()
    returns_total = filtered_data["Returns"].sum()

    # These ratios are descriptive indicators, not proof of profit.
    discount_rate = (
        discounts_total / gross_sales * 100
        if gross_sales != 0
        else None
    )

    return_rate = (
        returns_total / gross_sales * 100
        if gross_sales != 0
        else None
    )

    w1, w2 = st.columns(2)

    with w1:

        if discount_rate is not None:

            st.metric(
                "Discounts / Gross Sales",
                f"{discount_rate:.2f}%",
            )

        else:
            st.metric(
                "Discounts / Gross Sales",
                "N/A",
            )

    with w2:

        if return_rate is not None:

            st.metric(
                "Returns / Gross Sales",
                f"{return_rate:.2f}%",
            )

        else:
            st.metric(
                "Returns / Gross Sales",
                "N/A",
            )

    st.caption(
        "These ratios use the recorded Discounts and Returns values "
        "divided by Gross Sales. Their interpretation depends on how "
        "the source system records returns and discounts. They are not "
        "profit margins, and the dashboard does not calculate profit "
        "because a cost-of-goods measure has not been established."
    )

    # Export executive summary
    st.divider()

    st.subheader("Export Executive Summary")

    executive_summary = pd.DataFrame(
        {
            "Metric": [
                "Net Sales",
                "Total Sales",
                "Unique Orders",
                "Average Net Sales per Order",
                "Gross Sales",
                "Discounts",
                "Returns",
                "Shipping",
                "Taxes",
            ],
            "Value": [
                exec_net_sales,
                exec_total_sales,
                exec_orders,
                exec_aov,
                filtered_data["Gross sales"].sum(),
                discounts_total,
                returns_total,
                filtered_data["Shipping"].sum(),
                filtered_data["Taxes"].sum(),
            ],
        }
    )

    st.download_button(
        "Download Executive Summary CSV",
        data=executive_summary.to_csv(index=False).encode("utf-8"),
        file_name="executive_summary.csv",
        mime="text/csv",
        key="download_executive_summary",
    )


# ============================================================
# FOOTER
# ============================================================
st.divider()

footer_left, footer_right = st.columns([3, 1])

with footer_left:
    st.caption(
        "Sales Intelligence Dashboard | Interactive analytics "
        "built with Python, Pandas, Plotly, and Streamlit."
    )

with footer_right:
    st.caption("Data-driven decision support")