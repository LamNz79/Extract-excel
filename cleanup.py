import re
import pandas as pd
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

INPUT_FILE = os.path.join(BASE_DIR, "output.xlsx")
OUTPUT_FILE = os.path.join(BASE_DIR, "output_clean.xlsx")

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

# Clean columns
df["muc_cong_diem"] = df["muc_cong_diem"].apply(clean_muc_cong_diem)
df["so_diem"] = df["so_diem"].apply(clean_so_diem)

# Flag cleanup problems (optional but useful)
df["cleanup_issue"] = df["muc_cong_diem"].isna() | df["so_diem"].isna()

# OPTIONAL: drop rows that failed cleanup
# df = df[df["cleanup_issue"] == False]

# Save clean output
df.to_excel(OUTPUT_FILE, index=False)

print("✅ Cleanup done (old columns removed)")
print(f"📄 Clean file saved at: {OUTPUT_FILE}")
