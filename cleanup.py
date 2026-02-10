import re
import pandas as pd
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SANITIZED_DIR = os.path.join(BASE_DIR, "_sanitized")

INPUT_FILE = os.path.join(SANITIZED_DIR, "output.xlsx")
OUTPUT_FILE = os.path.join(SANITIZED_DIR, "output_clean.xlsx")

# =========================
# CLEANING FUNCTIONS
# =========================


def clean_muc_cong_diem(value):
    if pd.isna(value):
        return None
    match = re.search(r"\d+", str(value))
    return int(match.group()) if match else None


def clean_so_diem(value):
    if pd.isna(value):
        return None
    match = re.search(r"\d+(\.\d+)?", str(value))
    if not match:
        return None
    number = match.group()
    return int(number) if number.isdigit() else float(number)


# =========================
# MAIN
# =========================

df = pd.read_excel(INPUT_FILE)

# Store original values for comparison
df["muc_cong_diem_original"] = df["muc_cong_diem"]
df["so_diem_original"] = df["so_diem"]

# Clean columns
df["muc_cong_diem"] = df["muc_cong_diem"].apply(clean_muc_cong_diem)
df["so_diem"] = df["so_diem"].apply(clean_so_diem)

# Flag cleanup problems with detailed reasons
df["cleanup_issue"] = False
df["cleanup_reason"] = None

for idx, row in df.iterrows():
    issues = []
    
    if pd.notna(row["muc_cong_diem_original"]) and pd.isna(row["muc_cong_diem"]):
        issues.append(f"Không parse được mục cộng điểm: '{row['muc_cong_diem_original']}'")
    
    if pd.notna(row["so_diem_original"]) and pd.isna(row["so_diem"]):
        issues.append(f"Không parse được số điểm: '{row['so_diem_original']}'")
    
    if issues:
        df.at[idx, "cleanup_issue"] = True
        df.at[idx, "cleanup_reason"] = "; ".join(issues)

# Drop temporary columns
df = df.drop(columns=["muc_cong_diem_original", "so_diem_original"])

# OPTIONAL: drop rows that failed cleanup
# df = df[df["cleanup_issue"] == False]

# Save clean output
df.to_excel(OUTPUT_FILE, index=False)

print("✅ Cleanup done (old columns removed)")
print(f"📄 Clean file saved at: {OUTPUT_FILE}")
