import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# ---------------------------------------
# 1. Load Dataset
# ---------------------------------------
df = pd.read_excel("Product-Sales-Region.xlsx")
df.columns = df.columns.str.strip()

# Convert date columns
df["Date"] = pd.to_datetime(df["Date"])
df["Month"] = df["Date"].dt.to_period("M").astype(str)

# ---------------------------------------
# 2. Interactive Line Chart (Hover Effects)
# ---------------------------------------
fig_line = px.line(
    df.sort_values("Date"),
    x="Date",
    y="TotalPrice",
    color="Region",
    title="Interactive Sales Trend by Region",
    markers=True,
    hover_data={
        "Product": True,
        "Quantity": True,
        "Discount": True,
        "Date": True
    }
)

fig_line.update_layout(
    xaxis_title="Date",
    yaxis_title="Total Price",
    hovermode="x unified"
)

fig_line.show()

# ---------------------------------------
# 3. Dropdown Menu – Sales by Metric
# ---------------------------------------
fig_dropdown = go.Figure()

fig_dropdown.add_trace(
    go.Bar(x=df["Region"], y=df["TotalPrice"], visible=True, name="Total Sales")
)
fig_dropdown.add_trace(
    go.Bar(x=df["Region"], y=df["Quantity"], visible=False, name="Quantity Sold")
)
fig_dropdown.add_trace(
    go.Bar(x=df["Region"], y=df["Discount"], visible=False, name="Discount")
)

fig_dropdown.update_layout(
    title="Sales Metrics by Region (Dropdown)",
    xaxis_title="Region",
    yaxis_title="Value",
    updatemenus=[
        dict(
            buttons=[
                dict(label="Total Sales", method="update",
                     args=[{"visible": [True, False, False]},
                           {"yaxis": {"title": "Total Sales"}}]),
                dict(label="Quantity", method="update",
                     args=[{"visible": [False, True, False]},
                           {"yaxis": {"title": "Quantity"}}]),
                dict(label="Discount", method="update",
                     args=[{"visible": [False, False, True]},
                           {"yaxis": {"title": "Discount"}}])
            ],
            direction="down",
            showactive=True
        )
    ]
)

fig_dropdown.show()

# ---------------------------------------
# 4. Animated Scatter Plot – Monthly Sales
# ---------------------------------------
fig_anim = px.scatter(
    df,
    x="Quantity",
    y="TotalPrice",
    size="UnitPrice",
    color="Region",
    animation_frame="Month",
    hover_name="Product",
    title="Monthly Sales Animation",
    size_max=40
)

fig_anim.update_layout(
    xaxis_title="Quantity",
    yaxis_title="Total Price",
    transition={"duration": 600}
)

fig_anim.show()
