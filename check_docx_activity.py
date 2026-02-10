import os

# Đường dẫn đến thư mục chứa dữ liệu
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, "input")
SANITIZED_DIR = os.path.join(BASE_DIR, "_sanitized")


def check_files_in_folders():
    if not os.path.exists(INPUT_DIR):
        print(f"❌ Không tìm thấy thư mục: {INPUT_DIR}")
        return

    excel_files = []
    word_files = []
    ignored_items = []

    # Duyệt đệ quy trong thư mục _sanitized
    for root, dirs, files in os.walk(INPUT_DIR):
        # 1. Loại bỏ các thư mục ẩn và venv ngay lập tức
        dirs[:] = [
            d
            for d in dirs
            if not d.startswith(".") and d not in ["venv", "__pycache__"]
        ]

        for file in files:
            # 2. Bỏ qua các file script hoặc file cấu hình nếu lỡ nằm trong này
            if (
                file.endswith(".py")
                or file == "requirements.txt"
                or file.startswith("output")
            ):
                continue

            path = os.path.join(root, file)
            rel_path = os.path.relpath(path, INPUT_DIR)
            file_lower = file.lower()

            # 3. Phân loại dựa trên đuôi file
            if file_lower.endswith((".xlsx", ".xls", ".csv")):
                excel_files.append(rel_path)
            elif file_lower.endswith((".docx", ".doc")):
                word_files.append(rel_path)
            else:
                ignored_items.append(rel_path)

    # --- IN BÁO CÁO TỔNG KẾT ---
    print(f"📂 Thư mục quét: {INPUT_DIR}")
    print("-" * 50)
    print("📊 TỔNG HỢP:")
    print(f"✅ Excel tìm thấy: {len(excel_files)} file")
    print(f"📝 Word tìm thấy : {len(word_files)} file")
    print(f"❓ Định dạng khác: {len(ignored_items)} file")
    print("-" * 50)

    if word_files:
        print("Danh sách file Word (Sẽ được xử lý bởi extract_from_docx):")
        for f in word_files:
            print(f"  + {f}")

    if ignored_items:
        print("\nCác file không thuộc Excel/Word (Bị bỏ qua):")
        for f in ignored_items:
            print(f"  - {f}")


if __name__ == "__main__":
    check_files_in_folders()
