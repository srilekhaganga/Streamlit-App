import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime

# Set page config for wide layout
st.set_page_config(page_title="SuperStore KPI Dashboard", layout="wide")

# ---- Load Data ----
@st.cache_data
def load_data():
    df = pd.read_excel("Sample - Superstore-1.xlsx", engine="openpyxl")
    df["Order Date"] = pd.to_datetime(df["Order Date"])
    return df

df_original = load_data()
# ---- Sidebar Filters ----
st.sidebar.title("Filters")

# Region Filter
all_regions = sorted(df_original["Region"].dropna().unique())
selected_region = st.sidebar.selectbox("Select Region", options=["All"] + all_regions)

# State Filter
all_states = sorted(df_original["State"].dropna().unique())
selected_state = st.sidebar.selectbox("Select State", options=["All"] + all_states)

# Category Filter
all_categories = sorted(df_original["Category"].dropna().unique())
selected_category = st.sidebar.selectbox("Select Category", options=["All"] + all_categories)

# Sub-Category Filter
all_subcats = sorted(df_original["Sub-Category"].dropna().unique())
selected_subcat = st.sidebar.selectbox("Select Sub-Category", options=["All"] + all_subcats)

# Customer Segment Filter
all_segments = sorted(df_original["Segment"].dropna().unique())
selected_segment = st.sidebar.selectbox("Select Customer Segment", options=["All"] + all_segments)

# Customer Filter
all_customers = sorted(df_original["Customer Name"].dropna().unique())
selected_customer = st.sidebar.selectbox("Select Customer", options=["All"] + all_customers)

# Filter data
df = df_original.copy()
if selected_region != "All":
    df = df[df["Region"] == selected_region]
if selected_state != "All":
    df = df[df["State"] == selected_state]
if selected_category != "All":
    df = df[df["Category"] == selected_category]
if selected_subcat != "All":
    df = df[df["Sub-Category"] == selected_subcat]
if selected_segment != "All":
    df = df[df["Segment"] == selected_segment]
if selected_customer != "All":
    df = df[df["Customer Name"] == selected_customer]

# ---- KPI Calculation ----
total_sales = df["Sales"].sum() / 1e6  # Convert to millions
total_quantity = df["Quantity"].sum()
total_profit = df["Profit"].sum() / 1e6  # Convert to millions
total_returns = df[df["Order ID"].isin(df_original[df_original["Order ID"].duplicated()]["Order ID"])]["Sales"].sum()
return_rate = (total_returns / df_original["Sales"].sum()) * 100 if df_original["Sales"].sum() > 0 else 0
margin_rate = (total_profit * 1e6 / df["Sales"].sum()) * 100 if df["Sales"].sum() != 0 else 0  # Percentage
avg_order_value = df["Sales"].sum() / df["Order ID"].nunique() if df["Order ID"].nunique() > 0 else 0

# ---- KPI Display ----
kpi_col1, kpi_col2, kpi_col3, kpi_col4, kpi_col5 = st.columns(5)
with kpi_col1:
    st.metric(label="Total Sales (in Millions)", value=f"${total_sales:,.2f}M")
with kpi_col2:
    st.metric(label="Total Profit (in Millions)", value=f"${total_profit:,.2f}M")
with kpi_col3:
    st.metric(label="Return Rate (%)", value=f"{return_rate:.2f}%")
with kpi_col4:
    st.metric(label="Margin Rate (%)", value=f"{(margin_rate):,.2f}%")
with kpi_col5:
    st.metric(label="Avg Order Value ($)", value=f"${avg_order_value:,.2f}")

# ---- Trend Comparison ----
st.subheader("Sales Trends Over Time")
daily_sales = df.groupby(df["Order Date"].dt.to_period("M"))["Sales"].sum().reset_index()
daily_sales["Order Date"] = daily_sales["Order Date"].astype(str)
fig = px.line(daily_sales, x="Order Date", y="Sales", title="Monthly Sales Trend")
st.plotly_chart(fig, use_container_width=True)

# ---- Enhanced Product Analysis ----
st.subheader("Top Products by Sales")
top_products = df.groupby("Product Name")["Sales"].sum().nlargest(10).reset_index()
fig_bar = px.bar(top_products, x="Sales", y="Product Name", orientation="h", title="Top 10 Products")
st.plotly_chart(fig_bar, use_container_width=True)

# ---- Regional Performance ----
st.subheader("Regional Sales Performance")
profit_by_region = df.groupby("Region")["Profit"].sum().reset_index()
fig_region = px.bar(profit_by_region, x="Region", y="Profit", title="Profit by Region", color="Profit", color_continuous_scale="Blues")
st.plotly_chart(fig_region, use_container_width=True)

# ---- Return Analysis ----
st.subheader("Profitability Impact of Returns")
return_reasons = df.groupby("Sub-Category")["Profit"].sum().sort_values().reset_index()
fig_return = px.bar(return_reasons, x="Profit", y="Sub-Category", orientation="h", title="Least Profitable Sub-Categories")
st.plotly_chart(fig_return, use_container_width=True)

# ---- Customer-Based Insights ----
st.subheader("Top Customers by Sales")
top_customers = df.groupby("Customer Name")["Sales"].sum().nlargest(10).reset_index()
fig_customers = px.bar(top_customers, x="Sales", y="Customer Name", orientation="h", title="Top 10 Customers by Sales")
st.plotly_chart(fig_customers, use_container_width=True)
