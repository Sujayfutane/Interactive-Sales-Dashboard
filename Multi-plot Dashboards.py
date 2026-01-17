import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

# ---------------------------------------
# 1. Load Dataset
# ---------------------------------------
df = pd.read_excel("Product-Sales-Region.xlsx")
df.columns = df.columns.str.strip()

# Convert date column
df["Date"] = pd.to_datetime(df["Date"])

# ---------------------------------------
# 2. Global Theme (Coordinated Styling)
# ---------------------------------------
sns.set_theme(
    style="whitegrid",
    palette="deep",
    font_scale=0.9
)

# ---------------------------------------
# 3. Create 2×2 Subplot Grid
# ---------------------------------------
fig, axes = plt.subplots(2, 2, figsize=(14, 10))
fig.suptitle("Sales Performance Dashboard", fontsize=16, fontweight="bold")

# ---------------------------------------
# Plot 1: Line Plot – Sales Over Time
# ---------------------------------------
sns.lineplot(
    data=df.sort_values("Date"),
    x="Date",
    y="TotalPrice",
    ax=axes[0, 0],
    linewidth=2
)
axes[0, 0].set_title("Total Sales Over Time")
axes[0, 0].set_xlabel("Date")
axes[0, 0].set_ylabel("Total Price")

# ---------------------------------------
# Plot 2: Bar Plot – Total Sales by Region
# ---------------------------------------
sns.barplot(
    data=df,
    x="Region",
    y="TotalPrice",
    estimator=sum,
    ax=axes[0, 1]
)
axes[0, 1].set_title("Total Sales by Region")
axes[0, 1].set_xlabel("Region")
axes[0, 1].set_ylabel("Total Sales")

# ---------------------------------------
# Plot 3: Box Plot – Sales Distribution by Product
# ---------------------------------------
sns.boxplot(
    data=df,
    x="Product",
    y="TotalPrice",
    ax=axes[1, 0]
)
axes[1, 0].set_title("Sales Distribution by Product")
axes[1, 0].set_xlabel("Product")
axes[1, 0].set_ylabel("Total Price")
axes[1, 0].tick_params(axis='x', rotation=45)

# ---------------------------------------
# Plot 4: Heatmap – Correlation Matrix
# ---------------------------------------
numeric_df = df.select_dtypes(include=["int64", "float64"])
corr = numeric_df.corr()

sns.heatmap(
    corr,
    annot=True,
    fmt=".2f",
    cmap="coolwarm",
    ax=axes[1, 1],
    cbar=False
)
axes[1, 1].set_title("Correlation Heatmap")

# ---------------------------------------
# 4. Layout Adjustment
# ---------------------------------------
plt.tight_layout(rect=[0, 0, 1, 0.96])
plt.show()
