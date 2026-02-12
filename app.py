"""
Python to Browser App - Training Tool
A simple Streamlit application that reads Excel data, processes it, and displays charts.
"""

import streamlit as st
import pandas as pd
import plotly.express as px
from pathlib import Path

# -----------------------------
# Page configuration
# -----------------------------
st.set_page_config(
    page_title="Excel to Chart App",
    page_icon="📊",
    layout="wide"
)

st.markdown("""
<style>

/* KPI cards */
div[data-testid="stMetric"] {
    background-color: #ffffff;
    border: 1px solid #e6e9ef;
    padding: 18px;
    border-radius: 14px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.05);
}

/* KPI label */
div[data-testid="stMetric"] label {
    font-size: 14px;
    color: #6b7280;
}

/* KPI value */
div[data-testid="stMetric"] div {
    font-size: 30px;
    font-weight: 700;
    color: #111827;
}

</style>
""", unsafe_allow_html=True)

# -----------------------------
# App title (no header emojis)
# -----------------------------
st.title("Sales Performance Dashboard")
st.markdown(
    "<div class='small-muted'>Executive Revenue Overview | Interactive Reporting</div>",
    unsafe_allow_html=True
)

# File path
excel_file = Path("data.xlsx")

# -----------------------------
# Function to load data
# -----------------------------
@st.cache_data
def load_data(file_path):
    """Load data from Excel file"""
    try:
        df = pd.read_excel(file_path)
        return df, None
    except FileNotFoundError:
        return None, "Error: data.xlsx not found. Please ensure the file exists in the project folder."
    except Exception as e:
        return None, f"Error loading file: {str(e)}"

# Reload button
col1, col2, col3 = st.columns([1, 1, 3])
with col1:
    if st.button("Reload Data", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

# Load the data
df, error = load_data(excel_file)
if error:
    st.error(error)
    st.stop()

    # --- Normalize column names (prevents KeyError like 'Date ') ---
df.columns = df.columns.astype(str).str.strip()

# Optional: make common variations consistent
rename_map = {}
for c in df.columns:
    if c.lower() == "date":
        rename_map[c] = "Date"
    elif c.lower() == "units":
        rename_map[c] = "Units"
    elif c.lower() in ["unitprice", "unit price"]:
        rename_map[c] = "UnitPrice"
    elif c.lower() == "product":
        rename_map[c] = "Product"
    elif c.lower() == "region":
        rename_map[c] = "Region"

df = df.rename(columns=rename_map)


# -----------------------------
# Standardize / clean types
# -----------------------------
df["Date"] = pd.to_datetime(df.get("Date"), errors="coerce")
df["Units"] = pd.to_numeric(df.get("Units"), errors="coerce")
df["UnitPrice"] = pd.to_numeric(df.get("UnitPrice"), errors="coerce")

# Always calculate Revenue (even if already exists)
df["Revenue"] = df["Units"] * df["UnitPrice"]

# Drop fully invalid rows early (keeps filters stable)
df = df.dropna(subset=["Date", "Product", "Region", "Units", "UnitPrice"])

# -----------------------------
# Filters (clean + safe)
# -----------------------------
st.markdown("<div class='section-gap'></div>", unsafe_allow_html=True)
st.subheader("Filters")

f1, f2, f3 = st.columns(3)

regions = sorted(df["Region"].dropna().unique())
products = sorted(df["Product"].dropna().unique())

with f1:
    selected_region = st.multiselect(
        "Region",
        regions,
        default=regions
    )

with f2:
    selected_product = st.multiselect(
        "Product",
        products,
        default=products
    )

with f3:
    min_date = df["Date"].min()
    max_date = df["Date"].max()

    # If Date column is empty or invalid, guard
    if pd.isna(min_date) or pd.isna(max_date):
        st.warning("No valid dates found in the dataset.")
        date_range = None
    else:
        date_range = st.date_input(
            "Date Range",
            value=[min_date.date(), max_date.date()]
        )

# Apply filters safely (handle empty selections)
df_filtered = df.copy()

if selected_region:
    df_filtered = df_filtered[df_filtered["Region"].isin(selected_region)]
else:
    df_filtered = df_filtered.iloc[0:0]  # empty

if selected_product:
    df_filtered = df_filtered[df_filtered["Product"].isin(selected_product)]
else:
    df_filtered = df_filtered.iloc[0:0]  # empty

if date_range and len(date_range) == 2:
    start_date = pd.to_datetime(date_range[0])
    end_date = pd.to_datetime(date_range[1]) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
    df_filtered = df_filtered[(df_filtered["Date"] >= start_date) & (df_filtered["Date"] <= end_date)]

# Ensure key values exist
df_clean = df_filtered.dropna(subset=["Revenue"])

# -----------------------------
# KPI Tiles
# -----------------------------
st.markdown("<div class='section-gap'></div>", unsafe_allow_html=True)


# --- KPI Calculations ---
total_revenue = df_clean["Revenue"].sum()
total_units = df_clean["Units"].sum()
avg_unit_price = df_clean["UnitPrice"].mean()

top_product_series = (
    df_clean.groupby("Product")["Revenue"]
    .sum()
    .sort_values(ascending=False)
)

top_product = top_product_series.index[0] if len(top_product_series) else "—"
top_product_revenue = top_product_series.iloc[0] if len(top_product_series) else 0


st.markdown("### Executive KPIs")

k1, k2, k3, k4 = st.columns(4, gap="large")

with k1:
    st.metric("Total Revenue", f"${total_revenue:,.0f}")

with k2:
    st.metric("Total Units Sold", f"{total_units:,.0f}")

with k3:
    st.metric("Average Unit Price", f"${avg_unit_price:,.0f}")

with k4:
    st.metric("Top Performing Product", top_product, f"${top_product_revenue:,.0f}")


# -----------------------------
# Charts
# -----------------------------
st.markdown("<div class='section-gap'></div>", unsafe_allow_html=True)
st.subheader("Revenue by Product")

if df_clean.empty:
    st.warning("No chart to display (filtered dataset is empty).")
else:
    px.defaults.template = "plotly_white"

    rev_by_product = (
        df_clean.groupby("Product", as_index=False)["Revenue"]
        .sum()
        .sort_values("Revenue", ascending=True)
    )

    fig1 = px.bar(
        rev_by_product,
        x="Revenue",
        y="Product",
        orientation="h",
        text_auto=".2s",
        color="Revenue",
        color_continuous_scale="Blues",
        title=None
    )
    fig1.update_layout(
        showlegend=False,
        margin=dict(l=10, r=10, t=10, b=10),
        coloraxis_showscale=False
    )
    fig1.update_traces(textposition="outside", cliponaxis=False)
    st.plotly_chart(fig1, use_container_width=True)

    st.subheader("Revenue by Region")

rev_by_region = (
    df_clean.groupby("Region", as_index=False)["Revenue"]
    .sum()
    .sort_values("Revenue", ascending=False)
)

fig2 = px.pie(
    rev_by_region,
    names="Region",
    values="Revenue",
    template="plotly_white"
)

# Force it to be a normal pie (not donut)
fig2.update_traces(
    hole=0,
    textposition="inside",
    textinfo="percent+label"
)

fig2.update_layout(
    showlegend=True
)

st.plotly_chart(fig2, use_container_width=True)

st.subheader("Daily Revenue Trend")

if not df_clean.empty:

    # Ensure Date exists and is datetime
    if "Date" in df_clean.columns:
        df_clean["Date"] = pd.to_datetime(df_clean["Date"], errors="coerce")

        daily_rev = (
            df_clean.groupby("Date", as_index=False)["Revenue"]
            .sum()
        )

        if not daily_rev.empty:
            fig3 = px.line(
                daily_rev,
                x="Date",
                y="Revenue",
                markers=True,
                template="plotly_white"
            )

            fig3.update_traces(line=dict(width=3))
            fig3.update_layout(showlegend=False)

            st.plotly_chart(fig3, use_container_width=True)
        else:
            st.info("No data available for selected filters.")
    else:
        st.info("Date column not available.")
else:
    st.info("No data available for selected filters.")


# -----------------------------
# Data preview (filtered + optional CSV download)
# -----------------------------
st.markdown("<div class='section-gap'></div>", unsafe_allow_html=True)
st.subheader("Data Preview")

c1, c2, c3 = st.columns([1, 1, 2])
with c1:
    st.metric("Rows (filtered)", len(df_clean))
with c2:
    st.metric("Columns", len(df.columns))

with c3:
    csv = df_clean.to_csv(index=False).encode("utf-8")
    st.download_button(
        "Download filtered data (CSV)",
        data=csv,
        file_name="filtered_data.csv",
        mime="text/csv",
        use_container_width=True
    )

st.dataframe(df_clean, use_container_width=True, height=260)

# Footer
st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: #6b7280;'>"
    "Built with Streamlit | Update data.xlsx and click Reload Data to refresh"
    "</div>",
    unsafe_allow_html=True
)
