import os
import re
import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, "input")
OUT_DIR = os.path.join(BASE_DIR, "_sanitized")
MAX_FOOTER_SCAN_ROWS = 20


# =========================
# SAFE EXCEL OPEN
# =========================
def open_excel_safely(path):
    for engine in ["openpyxl", "xlrd", "odf"]:
        try:
            return pd.ExcelFile(path, engine=engine)
        except Exception:
            continue
    return None


# =========================
# STUDENT DETECTION (STRONG)
# =========================
def looks_like_student_row(row_cells):
    cells = [str(c).strip() for c in row_cells if str(c).strip()]
    joined = " ".join(cells)

    if re.search(r"\b\d{8,}\b", joined):
        return True

    if sum(1 for c in cells if c.isdigit()) >= 2:
        return True

    if len(cells) >= 3:
        return True

    return False


# =========================
# FOOTER DETECTION (WEAK)
# =========================
def looks_like_footer_row(row_text):
    text = row_text.strip()

    if not text:
        return True
    if "(" in text and ")" in text:
        return True
    if len(text.split()) <= 4:
        return True
    
    # Check for leadership/signature keywords
    if re.search(r"\b(TM\.|BCH|Chủ tịch|Phó|LÃNH ĐẠO|NGƯỜI LẬP|NGƯỜI SOẠN|TRƯỞNG KHOA|TRƯỞNG PHÒNG)\b", text, re.IGNORECASE):
        return True
    
    # Check for academic titles (TS., ThS., GS., PGS.)
    if re.search(r"\b(TS\.|ThS\.|GS\.|PGS\.|CN\.|KS\.)\s*[A-ZĐ]", text):
        return True
    
    # Names with 2-8 words (allowing multiple names in one row)
    if re.fullmatch(r"([A-ZĐ][a-zà-ỹ]+[\s]*){2,8}", text):
        return True

    return False


# =========================
# CLEAN ONE FILE (SAFE)
# =========================
def clean_excel_file(src_path, dst_path):
    xls = open_excel_safely(src_path)
    if xls is None:
        print(f"⚠️  Skipped (unreadable): {os.path.basename(src_path)}")
        return

    os.makedirs(os.path.dirname(dst_path), exist_ok=True)

    writer = pd.ExcelWriter(dst_path, engine="openpyxl")
    wrote_any_sheet = False

    for sheet in xls.sheet_names:
        try:
            df = xls.parse(sheet, header=None, dtype=str)
        except Exception:
            print(f"⚠️  Skipped sheet '{sheet}' in {os.path.basename(src_path)}")
            continue

        total_rows = len(df)
        footer_start = max(0, total_rows - MAX_FOOTER_SCAN_ROWS)
        keep_rows = []

        for i in range(total_rows):
            row_cells = df.iloc[i].fillna("").astype(str).tolist()
            joined = " ".join(row_cells)

            if i < footer_start:
                keep_rows.append(i)
            elif looks_like_student_row(row_cells):
                keep_rows.append(i)
            elif looks_like_footer_row(joined):
                continue
            else:
                keep_rows.append(i)

        df.iloc[keep_rows].to_excel(writer, sheet_name=sheet, index=False, header=False)
        wrote_any_sheet = True

    if wrote_any_sheet:
        writer.close()
        print(f"✅ Sanitized → {dst_path}")
    else:
        writer.close()
        os.remove(dst_path)
        print(f"⚠️  Nothing written: {src_path}")


# =========================
# PROCESS ALL FILES
# =========================
def process_all_files():
    # Tạo thư mục input nếu chưa có
    if not os.path.exists(INPUT_DIR):
        os.makedirs(INPUT_DIR)
        print(f"📁 Đã tạo thư mục: {INPUT_DIR}")
        print(f"⚠️  Vui lòng copy file Excel cần xử lý vào thư mục 'input'")
        return
    
    # Quét file trong thư mục input
    file_count = 0
    for root, _, files in os.walk(INPUT_DIR):
        for file in files:
            if not file.lower().endswith((".xlsx", ".xls", ".ods")):
                continue
            if file.startswith("output") or file.startswith("~$"):
                continue

            src = os.path.join(root, file)
            rel = os.path.relpath(src, INPUT_DIR)

            # 🔑 ALWAYS WRITE XLSX
            base_name = os.path.splitext(rel)[0]
            dst = os.path.join(OUT_DIR, base_name + ".xlsx")

            print(f"🧹 Sanitizing: {rel}")
            clean_excel_file(src, dst)
            file_count += 1
    
    if file_count == 0:
        print(f"⚠️  Không tìm thấy file Excel nào trong thư mục 'input'")
    else:
        print(f"\n📊 Đã xử lý: {file_count} file")


if __name__ == "__main__":
    process_all_files()
    print("✅ Footer sanitation finished (originals untouched)")
