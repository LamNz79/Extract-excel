import os
import pandas as pd
import warnings

# 1. Khử cảnh báo từ openpyxl (cho .xlsx) và các cảnh báo định dạng cũ
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, "input//UPLOAD PM//ĐOÀN TN (OK)")
print(INPUT_DIR)
KEY_PHRASE = "tên hoạt động:"


def detect_file(path):
    try:
        # Pandas sẽ tự động dùng engine 'xlrd' cho .xls và 'openpyxl' cho .xlsx
        xls = pd.ExcelFile(path)
    except Exception:
        # Nếu thiếu thư viện xlrd, nó sẽ báo lỗi ở đây
        return "SKIP", []

    hits = []
    for sheet in xls.sheet_names:
        try:
            df = xls.parse(sheet, header=None, dtype=str)
        except Exception:
            continue

        for i in range(len(df)):
            # Chuyển row thành list string để search key phrase
            row = df.iloc[i].fillna("").astype(str).tolist()
            joined = " ".join(row).strip().lower()

            if KEY_PHRASE in joined:
                hits.append(
                    {
                        "sheet": sheet,
                        "row": i + 1,
                        "text": joined[:80],
                    }
                )

    if len(hits) >= 2:
        return "REVIEW", hits
    return "OK", hits


def main():
    if not os.path.exists(INPUT_DIR):
        print(f"❌ Thư mục đầu vào không tồn tại: {INPUT_DIR}")
        return

    print(f"🔎 Scanning: {INPUT_DIR}\n")

    stats = {"excel": 0, "word": 0, "other": 0, "review": 0}
    others_list = []

    for root, dirs, files in os.walk(INPUT_DIR):
        # Loại bỏ thư mục rác/hệ thống
        dirs[:] = [
            d
            for d in dirs
            if not d.startswith(".") and d not in ["venv", "__pycache__"]
        ]

        for file in files:
            file_lower = file.lower()
            path = os.path.join(root, file)
            rel = os.path.relpath(path, INPUT_DIR)

            # Bỏ qua các file script và output
            if (
                file_lower.endswith(".py")
                or file == "requirements.txt"
                or file.startswith("output")
            ):
                continue

            # --- CẬP NHẬT: Kiểm tra cả .xlsx và .xls ---
            if file_lower.endswith((".xlsx", ".xls")):
                stats["excel"] += 1
                status, hits = detect_file(path)
                if status == "REVIEW":
                    stats["review"] += 1
                    print(f"⚠️  REVIEW (Multi-activity): {rel}")
                    for h in hits:
                        print(
                            f"    - Sheet '{h['sheet']}', row {h['row']}: {h['text']}"
                        )

            elif file_lower.endswith((".docx", ".doc")):
                stats["word"] += 1

            else:
                stats["other"] += 1
                others_list.append(rel)

    # --- BÁO CÁO TỔNG KẾT ---
    print("\n" + "=" * 40)
    print("📊 TỔNG HỢP FILE TRONG THƯ MỤC:")
    print(f"✅ File Excel (.xlsx, .xls): {stats['excel']}")
    print(f"📝 File Word (.docx, .doc): {stats['word']} (Đã bỏ qua)")
    print(f"❓ File khác:               {stats['other']}")
    print(
        f"🚨 Cần kiểm tra lại:        {stats['review']} file Excel có nhiều hoạt động"
    )

    if others_list:
        print("\nDanh sách file lạ không phải Excel/Word:")
        for item in others_list:
            print(f"  - {item}")
    print("=" * 40)


if __name__ == "__main__":
    main()
