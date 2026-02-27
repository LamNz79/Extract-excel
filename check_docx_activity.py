import os

# Base directory of the current Python file
# This ensures paths work correctly no matter where the script is run from
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Directory that contains the input data to be scanned
INPUT_DIR = os.path.join(BASE_DIR, "input")

# Directory intended to store sanitized/processed output (not used yet in this script)
SANITIZED_DIR = os.path.join(BASE_DIR, "_sanitized")


def check_files_in_folders():
    """
    Scan the INPUT_DIR directory recursively and classify files into:
    - Excel files (.xlsx, .xls, .csv)
    - Word files (.docx, .doc)
    - Other/unsupported file types (ignored)

    The function:
    - Skips hidden folders, virtual environments, and cache directories
    - Ignores script/config files if they appear in the input folder
    - Prints a summary report and detailed file lists
    """

    # Check if the input directory exists before proceeding
    if not os.path.exists(INPUT_DIR):
        print(f"❌ Không tìm thấy thư mục: {INPUT_DIR}")
        return

    # Containers for categorized files
    excel_files = []  # Stores relative paths to Excel files
    word_files = []  # Stores relative paths to Word files
    ignored_items = []  # Stores files with unsupported extensions

    # Walk through the INPUT_DIR recursively
    for root, dirs, files in os.walk(INPUT_DIR):
        # 1. Immediately remove unwanted directories from traversal
        #    - Hidden directories (starting with ".")
        #    - Virtual environment folders
        #    - Python cache folders
        dirs[:] = [
            d
            for d in dirs
            if not d.startswith(".") and d not in ["venv", "__pycache__"]
        ]

        # Process files in the current directory
        for file in files:
            # 2. Skip script files and configuration files
            #    These should not be treated as data inputs
            if (
                file.endswith(".py")
                or file == "requirements.txt"
                or file.startswith("output")
            ):
                continue

            # Full absolute path to the file
            path = os.path.join(root, file)

            # Relative path (used for cleaner reporting)
            rel_path = os.path.relpath(path, INPUT_DIR)

            # Lowercase filename for case-insensitive extension checks
            file_lower = file.lower()

            # 3. Classify files based on extension
            if file_lower.endswith((".xlsx", ".xls", ".csv")):
                excel_files.append(rel_path)
            elif file_lower.endswith((".docx", ".doc")):
                word_files.append(rel_path)
            else:
                # Files that are not Excel or Word are collected but ignored
                ignored_items.append(rel_path)

    # --- SUMMARY REPORT ---
    print(f"📂 Thư mục quét: {INPUT_DIR}")
    print("-" * 50)
    print("📊 TỔNG HỢP:")
    print(f"✅ Excel tìm thấy: {len(excel_files)} file")
    print(f"📝 Word tìm thấy : {len(word_files)} file")
    print(f"❓ Định dạng khác: {len(ignored_items)} file")
    print("-" * 50)

    # Detailed list of Word files (expected to be processed later)
    if word_files:
        print("Danh sách file Word (Sẽ được xử lý bởi extract_from_docx):")
        for f in word_files:
            print(f"  + {f}")

    # Detailed list of ignored files
    if ignored_items:
        print("\nCác file không thuộc Excel/Word (Bị bỏ qua):")
        for f in ignored_items:
            print(f"  - {f}")


# Entry point of the script
# Ensures the function runs only when this file is executed directly
if __name__ == "__main__":
    check_files_in_folders()
