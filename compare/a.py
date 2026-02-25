import json
from collections import Counter

# ===== Load file =====
with open("payload.json", "r", encoding="utf-8") as f:
    source_data = json.load(f)

with open("response.json", "r", encoding="utf-8") as f:
    target_data = json.load(f)

# ===== Lấy danh sách studentCode =====
source_codes = [item["studentCode"] for item in source_data]
target_codes = {item["studentCode"] for item in target_data.get("data", [])}

# ===== 1. Check trùng mã sinh viên trong file nguồn =====
counter = Counter(source_codes)
duplicate_codes = {code: count for code, count in counter.items() if count > 1}

# ===== 2. Tìm sinh viên có trong source nhưng thiếu trong target =====
missing_in_target = [code for code in source_codes if code not in target_codes]

# ===== In kết quả =====
print("===== KIỂM TRA TRÙNG MÃ SINH VIÊN (FILE NGUỒN) =====")
if duplicate_codes:
    for code, count in duplicate_codes.items():
        print(f"❌ {code} bị trùng {count} lần")
else:
    print("✅ Không có mã sinh viên bị trùng")

print("\n===== SINH VIÊN CÓ TRONG PAYLOAD NGUỒN NHƯNG THIẾU TRONG RESPONSE =====")
if missing_in_target:
    for code in missing_in_target:
        print(f"⚠️ Thiếu: {code}")
else:
    print("✅ Không thiếu sinh viên nào")
