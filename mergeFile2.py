import pandas as pd
import re

# Load the two CSV files
CSV_1 = r"C:\Users\aphiwut.j\Documents\Learning\resulting processing\2025-11-26T1522_Marks-GHC_25-T5_SIT40521_G-CKM.csv" #Files from CANVAS/ GOALS
CSV_2 = r"C:\Users\aphiwut.j\Documents\Learning\resulting processing\Book1.csv" #Files from Stars

df1 = pd.read_csv(CSV_1)
df2 = pd.read_csv(CSV_2)

def clean_id(series):
    return (
        series
        .astype(str)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)     # remove .0
        .str.replace(r"\s+", "", regex=True)      # remove spaces
        .str.replace(r"[^\d]", "", regex=True)    # keep only numbers
    )

# --- Clean column names ---
df1.columns = df1.columns.str.strip()
df2.columns = df2.columns.str.strip()

# --- Standardize ID columns ---
df1["SIS User ID"] = df1["SIS User ID"].astype(str).str.strip().str.lower()
df2["Student No"] = df2["Student No"].astype(str).str.strip().str.lower()

df1["SIS User ID"] = clean_id(df1["SIS User ID"])
df2["Student No"] = clean_id(df2["Student No"])

# --- Rename Student No to match ---
df2 = df2.rename(columns={"Student No": "SIS User ID"})
print(df2.head())

# --- Merge ---
merged_df = pd.merge(
    df1,
    df2,
    on="SIS User ID",
    how="inner"
)
print("merged_df: \n", merged_df)

# =====================================================
#                 NEW FILTERING LOGIC
# =====================================================

# --- Base columns to keep ---
base_columns = [
    "Enrol No",
    "SIS User ID",
    "SIS Login ID",
    "Section"
]

# --- Keep columns with "Final Score" but exclude the unwanted ones ---
final_score_columns = [
    col for col in merged_df.columns
    if "final score" in col.lower()
    and "unposted" not in col.lower()
    and "roll call" not in col.lower()
    and col.lower() != "final score"
]

# --- Combine final list of columns ---
output_columns = base_columns + final_score_columns

# --- Safety check ---
missing_cols = [col for col in base_columns if col not in merged_df.columns]
if missing_cols:
    raise Exception(f"Missing required columns: {missing_cols}")

# --- Select only those columns ---
final_output = merged_df[output_columns]

# --- Save output ---
final_output.to_csv("merged_final_scores_2.csv", index=False)

print("✅ Merge completed successfully: merged_final_scores_2.csv")
