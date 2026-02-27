import os
import pandas as pd
import warnings

# 1. Khử cảnh báo từ openpyxl (cho .xlsx) và các cảnh báo định dạng cũ
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, "input")

KEY_PHRASE = "tên hoạt động:"


def detect_file(path):
    try:
        # Pandas sẽ tự động dùng engine 'xlrd' cho .xls và 'openpyxl' cho .xlsx
        xls = pd.ExcelFile(path)
    except Exception:
        # Nếu thiếu thư viện xlrd, nó sẽ báo lỗi ở đây
        return "SKIP", [], 0, []

    num_sheets = len(xls.sheet_names)
    all_hits = []  # Tất cả các hoạt động tìm thấy
    sheets_with_multi_activity = []  # Sheets có >= 2 hoạt động
    
    for sheet in xls.sheet_names:
        try:
            df = xls.parse(sheet, header=None, dtype=str)
        except Exception:
            continue

        sheet_hits = []  # Hoạt động trong sheet này
        
        for i in range(len(df)):
            # Chuyển row thành list string để search key phrase
            row = df.iloc[i].fillna("").astype(str).tolist()
            joined = " ".join(row).strip().lower()

            if KEY_PHRASE in joined:
                hit = {
                    "sheet": sheet,
                    "row": i + 1,
                    "text": joined[:80],
                }
                sheet_hits.append(hit)
                all_hits.append(hit)
        
        # Kiểm tra sheet này có >= 2 hoạt động không
        if len(sheet_hits) >= 2:
            sheets_with_multi_activity.append(sheet)

    # Báo REVIEW nếu:
    # 1. Có >= 2 sheets
    # 2. Có sheet nào có >= 2 hoạt động
    has_multi_sheet = num_sheets >= 2
    has_sheet_with_multi_activity = len(sheets_with_multi_activity) > 0
    
    if has_multi_sheet or has_sheet_with_multi_activity:
        return "REVIEW", all_hits, num_sheets, sheets_with_multi_activity
    return "OK", all_hits, num_sheets, sheets_with_multi_activity


def main():
    if not os.path.exists(INPUT_DIR):
        print(f"❌ Thư mục đầu vào không tồn tại: {INPUT_DIR}")
        return

    print(f"🔎 Scanning: {INPUT_DIR}\n")

    stats = {
        "excel": 0,
        "word": 0,
        "other": 0,
        "review": 0,
        "multi_sheet": 0,
        "sheet_multi_activity": 0,
    }
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
                status, hits, num_sheets, sheets_with_multi_activity = detect_file(path)
                
                if status == "REVIEW":
                    stats["review"] += 1
                    
                    # Phân loại lý do REVIEW
                    reasons = []
                    if num_sheets >= 2:
                        stats["multi_sheet"] += 1
                        reasons.append(f"{num_sheets} sheets")
                    if sheets_with_multi_activity:
                        stats["sheet_multi_activity"] += 1
                        reasons.append(
                            f"Sheet có nhiều hoạt động: {', '.join(sheets_with_multi_activity)}"
                        )
                    
                    reason_text = " | ".join(reasons)
                    print(f"⚠️  REVIEW ({reason_text}): {rel}")
                    
                    if hits:
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
    print("\n" + "=" * 50)
    print("📊 TỔNG HỢP FILE TRONG THƯ MỤC:")
    print(f"✅ File Excel (.xlsx, .xls): {stats['excel']}")
    print(f"📝 File Word (.docx, .doc): {stats['word']} (Đã bỏ qua)")
    print(f"❓ File khác:               {stats['other']}")
    print(f"\n🚨 Cần kiểm tra lại:        {stats['review']} files")
    print(f"   📑 File nhiều sheets (>=2):         {stats['multi_sheet']} files")
    print(f"   📄 File có sheet chứa nhiều hoạt động: {stats['sheet_multi_activity']} files")

    if others_list:
        print("\nDanh sách file lạ không phải Excel/Word:")
        for item in others_list:
            print(f"  - {item}")
    print("=" * 50)


if __name__ == "__main__":
    main()
