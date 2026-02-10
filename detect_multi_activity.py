import os
import pandas as pd
import warnings

# 1. Khử triệt để UserWarning từ openpyxl (cái lỗi date bạn gặp ở đầu log)
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# Chỉ nên quét trong thư mục dữ liệu để tránh quét nhầm file code/git
SANITIZED_DIR = os.path.join(BASE_DIR, "_sanitized")

KEY_PHRASE = "tên hoạt động:"


def detect_file(path):
    try:
        xls = pd.ExcelFile(path)
    except Exception:
        return "SKIP", []

    hits = []
    for sheet in xls.sheet_names:
        try:
            df = xls.parse(sheet, header=None, dtype=str)
        except Exception:
            continue

        for i in range(len(df)):
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
    # Kiểm tra thư mục dữ liệu trước khi quét
    scan_target = BASE_DIR if os.path.exists(BASE_DIR) else BASE_DIR
    print(f"🔎 Scanning: {scan_target}\n")

    stats = {"excel": 0, "word": 0, "other": 0, "review": 0}
    others_list = []

    for root, dirs, files in os.walk(scan_target):
        # 2. LOẠI BỎ THƯ MỤC RÁC: .git, .vscode, venv
        dirs[:] = [
            d
            for d in dirs
            if not d.startswith(".") and d not in ["venv", "__pycache__"]
        ]

        for file in files:
            file_lower = file.lower()
            path = os.path.join(root, file)
            rel = os.path.relpath(path, BASE_DIR)

            # 3. Bỏ qua các file script của chính bạn
            if (
                file_lower.endswith(".py")
                or file == "requirements.txt"
                or file.startswith("output")
            ):
                continue

            # Phân loại và xử lý
            if file_lower.endswith(".xlsx"):
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
    print(f"✅ File Excel (.xlsx): {stats['excel']}")
    print(f"📝 File Word (.docx):  {stats['word']} (Đã bỏ qua trong bước này)")
    print(f"❓ File khác:          {stats['other']}")
    print(f"🚨 Cần kiểm tra lại:   {stats['review']} file Excel có nhiều hoạt động")

    if others_list:
        print("\nDanh sách file lạ không phải Excel/Word:")
        for item in others_list:
            print(f"  - {item}")
    print("=" * 40)


if __name__ == "__main__":
    main()
