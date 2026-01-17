import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from statsmodels.tsa.holtwinters import ExponentialSmoothing
import plotly.io as pio
import zipfile

# -----------------------------
# PAGE CONFIG
# -----------------------------
st.set_page_config(
    page_title="🚀 Advanced Sales Dashboard",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("🚀 Advanced Interactive Sales Dashboard")
st.markdown("Analyze your sales, track performance, and forecast future trends in one place.")

# -----------------------------
# LOAD DATA
# -----------------------------
df = pd.read_excel("Product-Sales-Region.xlsx")
df.columns = df.columns.str.strip()
df["Date"] = pd.to_datetime(df["Date"])
df["Month"] = df["Date"].dt.to_period("M").astype(str)

# -----------------------------
# SIDEBAR FILTERS
# -----------------------------
st.sidebar.header("Filters 🔎")

region_filter = st.sidebar.multiselect("Select Region", df["Region"].unique(), df["Region"].unique())
product_filter = st.sidebar.multiselect("Select Product", df["Product"].unique(), df["Product"].unique())
start_date, end_date = st.sidebar.date_input("Select Date Range", [df["Date"].min(), df["Date"].max()])

filtered_df = df[
    (df["Region"].isin(region_filter)) &
    (df["Product"].isin(product_filter)) &
    (df["Date"] >= pd.to_datetime(start_date)) &
    (df["Date"] <= pd.to_datetime(end_date))
]

# -----------------------------
# KPI METRICS
# -----------------------------
total_sales = filtered_df["TotalPrice"].sum()
total_orders = filtered_df["OrderID"].nunique()
avg_discount = filtered_df["Discount"].mean()

# Optional: calculate delta for KPIs
prev_df = df[
    (df["Region"].isin(region_filter)) &
    (df["Product"].isin(product_filter)) &
    (df["Date"] < pd.to_datetime(start_date))
]
prev_sales = prev_df["TotalPrice"].sum() if not prev_df.empty else 0
sales_delta = ((total_sales - prev_sales) / prev_sales * 100) if prev_sales != 0 else 0

col1, col2, col3 = st.columns(3)
col1.metric("💰 Total Sales", f"₹{total_sales:,.0f}", f"{sales_delta:+.2f}% vs previous period")
col2.metric("🧾 Total Orders", total_orders)
col3.metric("🏷 Avg Discount", f"{avg_discount:.2f}")

# -----------------------------
# TABS FOR BETTER LAYOUT
# -----------------------------
tabs = st.tabs(["📈 Sales Trend", "📊 Metrics by Region", "📦 Product Distribution", "🔮 Forecasting"])

# -----------------------------
# SALES TREND TAB
# -----------------------------
with tabs[0]:
    st.subheader("Sales Trend Over Time")
    trend_df = filtered_df.groupby("Date")["TotalPrice"].sum().reset_index()
    trend_df["Rolling_Avg"] = trend_df["TotalPrice"].rolling(7, min_periods=1).mean()
    trend_fig = px.line(trend_df, x="Date", y="TotalPrice", title="Daily Sales", markers=True, template="plotly_dark")
    trend_fig.add_scatter(x=trend_df["Date"], y=trend_df["Rolling_Avg"], mode="lines", name="7-Day Avg", line=dict(dash="dash"))
    st.plotly_chart(trend_fig, use_container_width=True)

# -----------------------------
# METRICS BY REGION TAB
# -----------------------------
with tabs[1]:
    st.subheader("Metrics by Region")
    metric = st.selectbox("Select Metric", ["TotalPrice", "Quantity", "Discount"])
    bar_fig = px.bar(filtered_df.groupby("Region")[metric].sum().reset_index(), x="Region", y=metric, color="Region",
                     title=f"{metric} by Region", template="plotly_dark")
    st.plotly_chart(bar_fig, use_container_width=True)

# -----------------------------
# PRODUCT DISTRIBUTION TAB
# -----------------------------
with tabs[2]:
    st.subheader("Sales Distribution by Product")
    col1, col2 = st.columns(2)
    with col1:
        box_fig = px.box(filtered_df, x="Product", y="TotalPrice", color="Product", template="plotly_dark",
                         title="Sales Distribution by Product")
        st.plotly_chart(box_fig, use_container_width=True)
    with col2:
        numeric_df = filtered_df.drop(columns=["OrderID"]).select_dtypes(include=["int64", "float64"])
        corr = numeric_df.corr()
        heatmap_fig = go.Figure(go.Heatmap(z=corr.values, x=corr.columns, y=corr.columns, colorscale="Viridis", zmid=0))
        heatmap_fig.update_layout(title="Correlation Heatmap", template="plotly_dark")
        st.plotly_chart(heatmap_fig, use_container_width=True)

# -----------------------------
# FORECASTING TAB
# -----------------------------
with tabs[3]:
    st.subheader("Sales Forecast (Next 6 Months)")
    monthly_sales = filtered_df.groupby("Month")["TotalPrice"].sum().reset_index()
    monthly_sales["Month"] = pd.to_datetime(monthly_sales["Month"])
    model = ExponentialSmoothing(monthly_sales["TotalPrice"], trend="add", seasonal=None).fit()
    forecast = model.forecast(6)
    forecast_df = pd.DataFrame({
        "Month": pd.date_range(start=monthly_sales["Month"].max() + pd.offsets.MonthBegin(1), periods=6, freq="MS"),
        "Forecast": forecast
    })
    forecast_fig = px.line(monthly_sales, x="Month", y="TotalPrice", title="Sales Forecast", markers=True, template="plotly_dark")
    forecast_fig.add_scatter(x=forecast_df["Month"], y=forecast_df["Forecast"], mode="lines+markers", name="Forecast",
                             line=dict(color="orange"), hovertemplate="Forecast: ₹%{y:,.0f}<br>Date: %{x|%b %Y}")
    st.plotly_chart(forecast_fig, use_container_width=True)

# -----------------------------
# EXPORT DASHBOARD
# -----------------------------
st.subheader("💾 Export Dashboard")
if st.button("Export Charts to ZIP"):
    files = ["trend.html", "region.html", "box.html", "heatmap.html", "forecast.html"]
    pio.write_html(trend_fig, "trend.html", auto_open=False)
    pio.write_html(bar_fig, "region.html", auto_open=False)
    pio.write_html(box_fig, "box.html", auto_open=False)
    pio.write_html(heatmap_fig, "heatmap.html", auto_open=False)
    pio.write_html(forecast_fig, "forecast.html", auto_open=False)
    with zipfile.ZipFile("dashboard.zip", 'w') as zipf:
        for file in files:
            zipf.write(file)
    st.success("✅ Dashboard exported as dashboard.zip")
