import os
import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
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
                        "row": i + 1,  # human-friendly
                        "text": joined[:80],
                    }
                )

    if len(hits) >= 2:
        return "REVIEW", hits

    return "OK", hits


def main():
    print("🔎 Checking for multiple activities via 'Tên hoạt động:'\n")

    for root, _, files in os.walk(BASE_DIR):
        for file in files:
            if not file.lower().endswith(".xlsx"):
                continue

            path = os.path.join(root, file)
            status, hits = detect_file(path)

            if status == "REVIEW":
                rel = os.path.relpath(path, BASE_DIR)
                print(f"⚠️  REVIEW: {rel}")
                for h in hits:
                    print(f"    - Sheet '{h['sheet']}', row {h['row']}: {h['text']}")


if __name__ == "__main__":
    main()
