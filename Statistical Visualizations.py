import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import numpy as np

# ---------------------------------------
# 1. Load Dataset
# ---------------------------------------
df = pd.read_excel("Product-Sales-Region.xlsx")

# Clean column names
df.columns = df.columns.str.strip()

# ---------------------------------------
# 2. Seaborn Theme
# ---------------------------------------
sns.set_theme(style="whitegrid", palette="Set2")

# ---------------------------------------
# 3. BOX PLOT – Total Sales by Region
# ---------------------------------------
plt.figure(figsize=(8, 4))
sns.boxplot(
    data=df,
    x="Region",
    y="TotalPrice",
    showmeans=True,
    meanprops={
        "marker": "o",
        "markerfacecolor": "red",
        "markeredgecolor": "black"
    }
)

plt.title("Total Sales Distribution by Region")
plt.xlabel("Region")
plt.ylabel("Total Price")
plt.show()

# ---------------------------------------
# 4. VIOLIN PLOT – Total Sales by Product
# ---------------------------------------
plt.figure(figsize=(9, 4))
sns.violinplot(
    data=df,
    x="Product",
    y="TotalPrice",
    inner="quartile",   # shows median & quartiles
    cut=0
)

plt.title("Sales Distribution by Product (Violin Plot)")
plt.xlabel("Product")
plt.ylabel("Total Price")
plt.show()

# ---------------------------------------
# 5. BOX + VIOLIN COMBINATION
# ---------------------------------------
plt.figure(figsize=(9, 4))
sns.violinplot(
    data=df,
    x="Region",
    y="TotalPrice",
    inner=None,
    alpha=0.6
)
sns.boxplot(
    data=df,
    x="Region",
    y="TotalPrice",
    width=0.2,
    showfliers=False
)

plt.title("Sales Distribution by Region (Violin + Box)")
plt.xlabel("Region")
plt.ylabel("Total Price")
plt.show()

# ---------------------------------------
# 6. STATISTICAL ANNOTATION – Mean & Median
# ---------------------------------------
plt.figure(figsize=(8, 4))
ax = sns.boxplot(data=df, x="Region", y="TotalPrice")

# Calculate statistics
stats = df.groupby("Region")["TotalPrice"].agg(["mean", "median"])

# Add annotations
for i, region in enumerate(stats.index):
    mean_val = stats.loc[region, "mean"]
    median_val = stats.loc[region, "median"]

    ax.text(
        i, mean_val,
        f"Mean: {mean_val:.0f}",
        ha="center",
        va="bottom",
        fontsize=9,
        color="blue"
    )

    ax.text(
        i, median_val,
        f"Median: {median_val:.0f}",
        ha="center",
        va="top",
        fontsize=9,
        color="darkred"
    )

plt.title("Sales Statistics by Region")
plt.xlabel("Region")
plt.ylabel("Total Price")
plt.show()
