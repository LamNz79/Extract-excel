import os
import re
import pandas as pd
from docx import Document
import warnings

# Silence openpyxl date warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# =========================
# BASE PATH
# =========================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, "input")
SANITIZED_DIR = os.path.join(BASE_DIR, "_sanitized")

# =========================
# FIELD DEFINITIONS
# =========================
FIELDS = {
    "muc_cong_diem": ["Mục cộng điểm"],
    "so_diem": ["Số điểm được cộng"],
    "ten_hoat_dong": ["Tên hoạt động"],
    "don_vi": ["Đơn vị", "Đơn vị tổ chức", "Đoàn Thanh niên", "Đoàn THANH NIÊN"],
}


# =========================
# HELPERS
# =========================
def clean_value(text):
    if ":" in text:
        return text.split(":", 1)[1].strip()
    return text.strip()


def looks_like_don_vi(text):
    text = text.strip()

    if len(text) < 5 or len(text) > 60:
        return False

    if not re.search(r"\b(KHOA|ĐOÀN|VIỆN|TRUNG TÂM|PHÒNG)\b", text):
        return False

    letters = [c for c in text if c.isalpha()]
    if not letters:
        return False

    uppercase_ratio = sum(1 for c in letters if c.isupper()) / len(letters)
    return uppercase_ratio > 0.6


# =========================
# EXCEL EXTRACTOR (MULTI-ACTIVITY)
# =========================
def extract_from_excel(path):
    xls = pd.ExcelFile(path, engine=None)

    activity = {
        "ten_hoat_dong": None,
        "muc_cong_diem": None,
        "so_diem": None,
        "don_vi": None,
    }

    for sheet in xls.sheet_names:
        df = xls.parse(sheet, header=None, dtype=str)

        for i in range(len(df)):
            for j in range(len(df.columns)):
                cell = str(df.iloc[i, j]).strip()

                if not cell or cell.lower() == "nan":
                    continue

                # ===== TEN HOAT DONG (FIRST ONE WINS) =====
                if activity["ten_hoat_dong"] is None and "Tên hoạt động" in cell:
                    if ":" in cell:
                        activity["ten_hoat_dong"] = clean_value(cell)
                    elif j + 1 < len(df.columns):
                        activity["ten_hoat_dong"] = str(df.iloc[i, j + 1]).strip()
                    continue

                # ===== DON VI =====
                if activity["don_vi"] is None and looks_like_don_vi(cell):
                    activity["don_vi"] = cell
                    continue

                # ===== MUC CONG DIEM =====
                if activity["muc_cong_diem"] is None and "Mục cộng điểm" in cell:
                    activity["muc_cong_diem"] = clean_value(cell)
                    continue

                # ===== SO DIEM =====
                if activity["so_diem"] is None and "Số điểm được cộng" in cell:
                    activity["so_diem"] = clean_value(cell)
                    continue

    return [activity]  # 👈 ALWAYS ONE ACTIVITY


# =========================
# DOCX EXTRACTOR (MULTI-ACTIVITY)
# =========================
def extract_from_docx(path):
    doc = Document(path)
    activities = []
    current = None

    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue

        if "Tên hoạt động" in text:
            if current:
                activities.append(current)

            current = {
                "ten_hoat_dong": clean_value(text),
                "muc_cong_diem": None,
                "so_diem": None,
                "don_vi": None,
            }
            continue

        if current is None:
            continue

        if "Mục cộng điểm" in text:
            current["muc_cong_diem"] = clean_value(text)

        elif "Số điểm được cộng" in text:
            current["so_diem"] = clean_value(text)

        elif current["don_vi"] is None and looks_like_don_vi(text):
            current["don_vi"] = text

    if current:
        activities.append(current)

    return activities


# =========================
# MAIN PROCESS (RECURSIVE)
# =========================
records = []

for root, dirs, files in os.walk(SANITIZED_DIR):
    for file in files:
        if file.endswith(".py") or file.startswith("output"):
            continue

        file_path = os.path.join(root, file)
        rel_path = os.path.relpath(root, SANITIZED_DIR)

        try:
            if file.endswith((".xlsx", ".xls", ".ods")):
                activities = extract_from_excel(file_path)
            elif file.endswith(".docx"):
                activities = extract_from_docx(file_path)
            else:
                continue

            for idx, act in enumerate(activities, start=1):
                row = {
                    "thu_muc": rel_path,
                    "source_file": file,
                    "activity_index": idx,
                    "ten_hoat_dong": act.get("ten_hoat_dong"),
                    "muc_cong_diem": act.get("muc_cong_diem"),
                    "so_diem": act.get("so_diem"),
                    "don_vi": act.get("don_vi"),
                    "needs_review": False,
                    "review_reason": None,
                }

                # Check for missing fields
                missing_fields = []
                for k in ["ten_hoat_dong", "muc_cong_diem", "so_diem", "don_vi"]:
                    if not row[k]:
                        missing_fields.append(k)
                
                if missing_fields:
                    row["needs_review"] = True
                    row["review_reason"] = f"Thiếu: {', '.join(missing_fields)}"

                records.append(row)

        except Exception as e:
            records.append(
                {
                    "thu_muc": rel_path,
                    "source_file": file,
                    "activity_index": None,
                    "ten_hoat_dong": None,
                    "muc_cong_diem": None,
                    "so_diem": None,
                    "don_vi": None,
                    "needs_review": True,
                    "review_reason": f"Lỗi xử lý file: {str(e)}",
                }
            )

# =========================
# OUTPUT
# =========================
output_path = os.path.join(SANITIZED_DIR, "output.xlsx")
df = pd.DataFrame(records)
df.to_excel(output_path, index=False)

print("✅ Extraction done")
print(f"📄 Output saved at: {output_path}")
