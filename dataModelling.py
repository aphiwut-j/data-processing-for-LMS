import pandas as pd

CSV_1 = r"C:\Users\aphiwut.j\Documents\Learning\resulting processing\Book1.csv" #Path to CSV from Stars, usually named Book1
CSV_2 = r"" #Path to CSV from Canvas/GOALS

# Load your merged CSV
# df = pd.read_csv(CSV_1)
df = pd.read_csv(
    CSV_1,
    keep_default_na=False,   # Do NOT convert default NA strings
    na_values=[]             # No custom NA values
)
df.columns = df.columns.str.strip().str.replace("\ufeff", "")

# print(df.columns.tolist())

# Base columns used everywhere
base_columns = [
    "Student No",
    "Enrol No",
    "First Name",
    "Last Name"
]

# ---------------- Sheet 1: Basic info ONLY (excluding on leave / finished) ----------------
sheet1 = df[base_columns + ["Status"]]

sheet1 = sheet1[
    ~sheet1["Status"].str.lower().isin(["on leave", "finished"])
]

# Now remove Status column (you don't want it shown)
sheet1 = sheet1[base_columns]

# ---------------- Sheet 2: Status filter (only on leave / finished) ----------------
sheet2_columns = base_columns + ["Status"]

sheet2 = df[sheet2_columns]

sheet2 = sheet2[
    sheet2["Status"].str.lower().isin(["on leave", "finished"])
]

# ---------------- Sheet 3: Outstanding fee >= 500 ----------------
sheet3_columns = base_columns + ["Outstanding Fees"]
sheet3 = df[sheet3_columns]

# Step 1: Replace "-" with 0
sheet3["Outstanding Fees"] = sheet3["Outstanding Fees"].replace("-", "0")

# Step 2: Clean money format (remove $ and commas)
sheet3["Outstanding Fees"] = (
    sheet3["Outstanding Fees"]
    .astype(str)
    .str.replace("$", "", regex=False)
    .str.replace(",", "", regex=False)
    .str.strip()
)

sheet3["Outstanding Fees"] = pd.to_numeric(sheet3["Outstanding Fees"], errors='coerce')
sheet3 = sheet3[
    sheet3["Outstanding Fees"] >= 500
]

# ---------------- Write all 3 sheets to one Excel file ----------------
with pd.ExcelWriter("student_report.xlsx", engine="openpyxl") as writer:
    sheet1.to_excel(writer, sheet_name="Basic_Info", index=False)
    sheet2.to_excel(writer, sheet_name="Status_Filtered", index=False)
    sheet3.to_excel(writer, sheet_name="Outstanding_500_Plus", index=False)

print("✅ student_report.xlsx created with your updated rules!")
