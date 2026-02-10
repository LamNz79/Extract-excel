import os

# Xác định thư mục gốc là nơi file script này đang đứng
BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def check_files_in_folders():
    print(f"🔍 Đang quét toàn bộ file tại: {BASE_DIR}\n")

    # Duyệt đệ quy qua tất cả các thư mục và file
    for root, _, files in os.walk(BASE_DIR):
        for file in files:
            # Bỏ qua chính file script đang chạy (nếu nó nằm cùng folder)
            if file == os.path.basename(__file__):
                continue

            path = os.path.join(root, file)
            file_lower = file.lower()

            # Kiểm tra nếu là file Word bị lẫn vào
            if file_lower.endswith((".doc", ".docx")):
                rel_path = os.path.relpath(path, BASE_DIR)
                print(f"❌ CẢNH BÁO: Phát hiện file Word: {rel_path}")
                continue

            # Kiểm tra nếu là file Excel thì mới xử lý tiếp
            if file_lower.endswith((".xlsx", ".xls", ".csv")):
                # Giả sử đây là hàm detect_file bạn đã có sẵn
                # status, hits = detect_file(path)

                # Logic xử lý Excel của bạn ở đây...
                rel_path = os.path.relpath(path, BASE_DIR)
                print(f"✅ Đang xử lý Excel: {rel_path}")

            else:
                # Các loại file khác (ảnh, pdf, v.v.) nếu cần bỏ qua
                pass


if __name__ == "__main__":
    check_files_in_folders()
