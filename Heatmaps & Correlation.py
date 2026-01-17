import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

# ---------------------------------------
# 1. Load Dataset
# ---------------------------------------
df = pd.read_excel("Product-Sales-Region.xlsx")

# Clean column names
df.columns = df.columns.str.strip()

print("Dataset Columns:")
print(df.columns)

# ---------------------------------------
# 2. Select Numerical Columns Only
# ---------------------------------------
numeric_df = df.select_dtypes(include=["int64", "float64"])

print("\nNumerical Columns Used for Correlation:")
print(numeric_df.columns)

# ---------------------------------------
# 3. Correlation Matrix
# ---------------------------------------
corr_matrix = numeric_df.corr()

print("\nCorrelation Matrix:")
print(corr_matrix)

# ---------------------------------------
# 4. Heatmap Visualization
# ---------------------------------------
plt.figure(figsize=(10, 6))
sns.heatmap(
    corr_matrix,
    annot=True,
    fmt=".2f",
    cmap="coolwarm",
    linewidths=0.5,
    linecolor="white",
    square=True,
    cbar_kws={"shrink": 0.8}
)

plt.title("Correlation Heatmap of Sales Dataset", fontsize=14)
plt.xticks(rotation=45, ha="right")
plt.yticks(rotation=0)
plt.tight_layout()
plt.show()
