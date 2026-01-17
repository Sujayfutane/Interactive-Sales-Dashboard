import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

# ---------------------------------------
# 1. Load Excel Dataset
# ---------------------------------------
df = pd.read_excel("Product-Sales-Region.xlsx")

# Clean column names
df.columns = df.columns.str.strip()

# Convert date columns
df["Date"] = pd.to_datetime(df["Date"])
df["OrderDate"] = pd.to_datetime(df["OrderDate"])

print("Columns:", df.columns.tolist())
print(df.head())

# ---------------------------------------
# 2. Seaborn Theme
# ---------------------------------------
sns.set_theme(style="whitegrid", palette="deep")

# ---------------------------------------
# 3. Line Plot – Total Sales Over Time
# ---------------------------------------
plt.figure(figsize=(9, 4))
sns.lineplot(
    data=df.sort_values("Date"),
    x="Date",
    y="TotalPrice",
    linewidth=2
)
plt.title("Total Sales Over Time")
plt.xlabel("Date")
plt.ylabel("Total Price")
plt.show()

# ---------------------------------------
# 4. Bar Plot – Total Sales by Region
# ---------------------------------------
plt.figure(figsize=(7, 4))
sns.barplot(
    data=df,
    x="Region",
    y="TotalPrice",
    estimator=sum,
    palette="viridis"
)
plt.title("Total Sales by Region")
plt.xlabel("Region")
plt.ylabel("Total Sales")
plt.show()

# ---------------------------------------
# 5. Scatter Plot – Quantity vs Total Price
# ---------------------------------------
plt.figure(figsize=(7, 4))
sns.scatterplot(
    data=df,
    x="Quantity",
    y="TotalPrice",
    hue="Region",
    alpha=0.7
)
plt.title("Quantity vs Total Price")
plt.xlabel("Quantity")
plt.ylabel("Total Price")
plt.show()

# ---------------------------------------
# 6. Histogram – Distribution of Total Sales
# ---------------------------------------
plt.figure(figsize=(7, 4))
sns.histplot(
    df["TotalPrice"],
    bins=30,
    kde=True,
    color="purple"
)
plt.title("Distribution of Total Sales")
plt.xlabel("Total Price")
plt.ylabel("Frequency")
plt.show()

# ---------------------------------------
# 7. Box Plot – Total Sales by Product
# ---------------------------------------
plt.figure(figsize=(9, 4))
sns.boxplot(
    data=df,
    x="Product",
    y="TotalPrice",
    palette="pastel"
)
plt.title("Sales Distribution by Product")
plt.xlabel("Product")
plt.ylabel("Total Price")
plt.show()
