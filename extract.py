import os
import re
import warnings

import pandas as pd
from docx import Document

# Silence warnings from openpyxl related to date parsing
# These warnings are harmless for this extraction use case
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# =========================
# BASE PATH
# =========================

# Absolute path of the current script
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Root directory containing raw input files (Excel / Word)
INPUT_DIR = os.path.join(BASE_DIR, "input")

# Directory for storing sanitized / extracted output files
SANITIZED_DIR = os.path.join(BASE_DIR, "_sanitized")

# =========================
# FIELD DEFINITIONS
# =========================

# Mapping of internal field keys to possible Vietnamese labels
# (Currently informational / future-proofing; not directly used in logic)
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
    """
    Clean a label-value string by removing the label part.

    Examples:
        "Tên hoạt động: ABC" -> "ABC"
        "ABC" -> "ABC"

    Args:
        text (str): Raw text extracted from a cell or paragraph

    Returns:
        str: Cleaned value without leading labels or extra whitespace
    """
    if ":" in text:
        return text.split(":", 1)[1].strip()
    return text.strip()


def looks_like_don_vi(text):
    """
    Heuristic check to determine whether a text string looks like
    a 'Đơn vị / Đơn vị tổ chức' value.

    The heuristic is based on:
    - Reasonable length
    - Presence of organizational keywords
    - Majority uppercase characters (common in official unit names)

    Args:
        text (str): Text candidate to evaluate

    Returns:
        bool: True if text likely represents an organization/unit
    """
    text = text.strip()

    # Reject strings that are too short or too long
    if len(text) < 5 or len(text) > 60:
        return False

    # Must contain common Vietnamese organization keywords
    if not re.search(r"\b(KHOA|ĐOÀN|VIỆN|TRUNG TÂM|PHÒNG)\b", text):
        return False

    # Extract alphabetic characters only
    letters = [c for c in text if c.isalpha()]
    if not letters:
        return False

    # Ratio of uppercase letters (used as a signal for official naming style)
    uppercase_ratio = sum(1 for c in letters if c.isupper()) / len(letters)
    return uppercase_ratio > 0.6


# =========================
# EXCEL EXTRACTOR (MULTI-ACTIVITY)
# =========================
def extract_from_excel(path):
    """
    Extract activity information from an Excel file.

    Assumptions:
    - Each Excel file represents ONE activity
    - Labels and values may appear in arbitrary positions
    - The first valid occurrence of each field is used

    Args:
        path (str): Absolute path to the Excel file

    Returns:
        list[dict]: A list containing exactly ONE activity dictionary
    """
    # Load Excel file with pandas (engine auto-detected)
    xls = pd.ExcelFile(path, engine=None)

    # Initialize activity container
    activity = {
        "ten_hoat_dong": None,
        "muc_cong_diem": None,
        "so_diem": None,
        "don_vi": None,
    }

    # Iterate through all sheets
    for sheet in xls.sheet_names:
        # Read sheet as raw text (no headers)
        df = xls.parse(sheet, header=None, dtype=str)

        # Traverse every cell in the sheet
        for i in range(len(df)):
            for j in range(len(df.columns)):
                cell = str(df.iloc[i, j]).strip()

                # Skip empty or NaN-like cells
                if not cell or cell.lower() == "nan":
                    continue

                # ===== TEN HOAT DONG (FIRST ONE WINS) =====
                # Capture activity name only once
                if activity["ten_hoat_dong"] is None and "Tên hoạt động" in cell:
                    if ":" in cell:
                        activity["ten_hoat_dong"] = clean_value(cell)
                    elif j + 1 < len(df.columns):
                        # Fallback: value in the adjacent cell
                        activity["ten_hoat_dong"] = str(df.iloc[i, j + 1]).strip()
                    continue

                # ===== DON VI =====
                # Use heuristic detection for organization/unit
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

    # Always return a list for consistency with DOCX extractor
    return [activity]  # 👈 ALWAYS ONE ACTIVITY


# =========================
# DOCX EXTRACTOR (MULTI-ACTIVITY)
# =========================
def extract_from_docx(path):
    """
    Extract one or more activities from a Word (.docx) document.

    Assumptions:
    - Each "Tên hoạt động" starts a new activity block
    - Subsequent related fields belong to the most recent activity

    Args:
        path (str): Absolute path to the Word file

    Returns:
        list[dict]: List of extracted activity dictionaries
    """
    doc = Document(path)
    activities = []
    current = None

    # Iterate through all paragraphs in the document
    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue

        # Detect start of a new activity
        if "Tên hoạt động" in text:
            # Save previous activity before starting a new one
            if current:
                activities.append(current)

            current = {
                "ten_hoat_dong": clean_value(text),
                "muc_cong_diem": None,
                "so_diem": None,
                "don_vi": None,
            }
            continue

        # Ignore content before the first activity
        if current is None:
            continue

        # Field extraction within an activity block
        if "Mục cộng điểm" in text:
            current["muc_cong_diem"] = clean_value(text)

        elif "Số điểm được cộng" in text:
            current["so_diem"] = clean_value(text)

        elif current["don_vi"] is None and looks_like_don_vi(text):
            current["don_vi"] = text

    # Append the last activity if present
    if current:
        activities.append(current)

    return activities


# =========================
# MAIN PROCESS (RECURSIVE)
# =========================

# Container for all extracted rows
records = []

# Walk through input directory recursively
for root, dirs, files in os.walk(INPUT_DIR):
    for file in files:
        # Skip scripts and generated output files
        if file.endswith(".py") or file.startswith("output"):
            continue

        file_path = os.path.join(root, file)
        rel_path = os.path.relpath(root, INPUT_DIR)

        try:
            # Choose extractor based on file type
            if file.endswith((".xlsx", ".xls", ".ods")):
                activities = extract_from_excel(file_path)
            elif file.endswith(".docx"):
                activities = extract_from_docx(file_path)
            else:
                continue

            # Convert each activity into a flat output row
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

                # Detect missing mandatory fields
                missing_fields = []
                for k in ["ten_hoat_dong", "muc_cong_diem", "so_diem", "don_vi"]:
                    if not row[k]:
                        missing_fields.append(k)

                # Flag records requiring manual review
                if missing_fields:
                    row["needs_review"] = True
                    row["review_reason"] = f"Thiếu: {', '.join(missing_fields)}"

                records.append(row)

        except Exception as e:
            # Capture file-level processing errors
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

# Final Excel output path
output_path = os.path.join(SANITIZED_DIR, "output.xlsx")

# Convert records to DataFrame and export
df = pd.DataFrame(records)
df.to_excel(output_path, index=False)

print("✅ Extraction done")
print(f"📄 Output saved at: {output_path}")
