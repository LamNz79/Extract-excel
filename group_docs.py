import os
import shutil
from pathlib import Path

# ─────────────────────────────────────────────
# Paths
# ─────────────────────────────────────────────
BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / "input"
GROUP_DIR = BASE_DIR / "output" / "grouped_docs"


def collect_word_files() -> list[Path]:
    files = []
    for root, dirs, filenames in os.walk(INPUT_DIR):
        dirs[:] = [
            d
            for d in dirs
            if not d.startswith(".") and d not in ("venv", "__pycache__")
        ]
        for f in filenames:
            if f.lower().endswith((".docx", ".doc")) and not f.startswith("output"):
                files.append(Path(root) / f)
    return sorted(files)


def group_docs():
    if not INPUT_DIR.exists():
        print(f"❌ Không tìm thấy thư mục: {INPUT_DIR}")
        return

    word_files = collect_word_files()

    if not word_files:
        print("⚠️  Không tìm thấy file Word nào.")
        return

    # Create output folder (clear if already exists)
    if GROUP_DIR.exists():
        shutil.rmtree(GROUP_DIR)
    GROUP_DIR.mkdir(parents=True)

    print(f"📂 Quét từ  : {INPUT_DIR}")
    print(f"📁 Lưu vào  : {GROUP_DIR}")
    print("-" * 50)

    copied = []
    skipped = []

    for src in word_files:
        dest = GROUP_DIR / src.name

        # Handle filename collisions by appending a counter
        if dest.exists():
            stem, suffix = src.stem, src.suffix
            counter = 1
            while dest.exists():
                dest = GROUP_DIR / f"{stem}_{counter}{suffix}"
                counter += 1
            skipped.append((src, dest))  # renamed, not truly skipped

        shutil.copy2(src, dest)
        copied.append((src, dest))
        print(f"  ✅ {src.relative_to(INPUT_DIR)}  →  {dest.name}")

    print("-" * 50)
    print(f"✅ Đã copy {len(copied)} file vào: {GROUP_DIR}")


if __name__ == "__main__":
    group_docs()
