# PIPELINE V2 - MIGRATION GUIDE

## 🚨 CRITICAL FIXES

### 1. Footer Removal Corruption (FIXED)
**Problem:** Original script was corrupting Excel files
**Root Causes:**
- Wrote files without preserving original structure
- No dry-run mode for safety testing
- No backup mechanism
- Unclear logging about what was being removed

**Solution in v2:**
✅ DRY_RUN mode (enabled by default)
✅ Automatic backups (.backup files)
✅ Detailed preview of what will be removed
✅ Better student data detection (never removes student rows)
✅ More conservative footer detection

### 2. Multiple Activities Per File (FIXED)
**Problem:** Excel extractor always returned exactly 1 activity
**Impact:** Lost data when multiple activities existed in one file

**Solution in v2:**
✅ Properly detects multiple "Tên hoạt động" occurrences
✅ Skips student data rows during extraction
✅ Reports files with multiple activities
✅ Compatible with DOCX multi-activity extraction

### 3. Data Quality Validation (NEW)
**Problem:** No validation after cleanup
**Impact:** Bad data could slip into system

**Solution in v2:**
✅ Validates numeric ranges (mục: 1-10, điểm: 0-100)
✅ Checks text field lengths
✅ Generates validation report
✅ Statistics dashboard
✅ Flags records needing review

---

## 📋 UPDATED PIPELINE WORKFLOW

### Step 1: Test Footer Removal (DRY RUN)
```bash
python remove_footer_v2.py
```
**What it does:**
- Scans all Excel files
- Shows what WOULD be removed
- NO files are modified
- Review the output carefully

**Check the output:**
- Look for any student rows being flagged for removal
- Verify footer detection is working correctly

### Step 2: Run Footer Removal (LIVE)
**Only after dry run looks good:**

1. Open `remove_footer_v2.py`
2. Change line: `DRY_RUN = False`
3. Run: `python remove_footer_v2.py`

**Safety:**
- Original files are backed up as `.backup`
- If something goes wrong, you can restore

### Step 3: Extract Activities
```bash
python extract_v2.py
```
**What's new:**
- Detects multiple activities per file
- Reports statistics
- Better don_vi detection (more keywords, lower threshold)
- Skips student data rows

**Output:** `output.xlsx`

### Step 4: Clean & Validate
```bash
python cleanup_v2.py
```
**What's new:**
- Cleans field values
- Validates data quality
- Generates statistics
- Creates validation report

**Output:**
- `output_clean.xlsx` - Clean data
- `validation_report.txt` - Issues to review

---

## 🔍 KEY IMPROVEMENTS

### Don Vi Detection (Enhanced)
**Old:** Required 60% uppercase, limited keywords
**New:**
- Lowered to 40% uppercase (handles mixed case better)
- Added keywords: BAN, HỘI, CLB, ĐẠI HỌC
- Increased max length to 100 chars

### Student Data Protection (Stronger)
**Detection signals:**
1. ✅ 8+ digit ID (MSSV)
2. ✅ 2+ numeric cells (STT + MSSV)
3. ✅ 3+ populated columns (data row pattern)

**Guarantee:** If ANY signal matches → row is NEVER deleted

### Error Handling (Improved)
- Multi-engine Excel opening (openpyxl → xlrd → odf)
- Graceful sheet-level failures (skip bad sheets, continue)
- Detailed error logging
- Statistics tracking

### Validation (New)
**Checks:**
- Mục cộng điểm: 1-10 range
- Số điểm: 0-100 range
- Tên hoạt động: 5-200 chars
- Đơn vị: 3+ chars

**Report includes:**
- Validation failures by file
- Distribution statistics
- Top organizations
- Score statistics

---

## 📊 EXPECTED OUTPUT

### extract_v2.py Statistics
```
📊 EXTRACTION STATISTICS
========================================
Total files processed: 127
  - Excel files: 98
  - DOCX files: 29
  - Skipped: 0
Total activities extracted: 142
Files with multiple activities: 15
Errors: 2
Records needing review: 8
========================================
```

### cleanup_v2.py Statistics
```
📊 STATISTICS
========================================
Total records: 142
Records needing review: 12
Validation failures: 8

📌 Mục cộng điểm distribution:
   Mục 1: 23 activities
   Mục 2: 45 activities
   Mục 3: 38 activities
   Mục 5: 36 activities

📌 Score statistics:
   Min: 1.0
   Max: 15.0
   Mean: 5.73
   Median: 5.0
========================================
```

---

## ⚠️ MIGRATION CHECKLIST

### Before Running v2:
- [ ] Backup your entire data folder
- [ ] Review the new scripts (understand changes)
- [ ] Run footer removal in DRY RUN mode first
- [ ] Check dry run output for issues

### After Running v2:
- [ ] Compare output.xlsx row count with v1
- [ ] Review validation_report.txt
- [ ] Check for files with multiple activities (expected)
- [ ] Verify student data wasn't removed
- [ ] Spot-check a few Excel files for footer removal quality

### If Issues Found:
- [ ] Restore from .backup files if needed
- [ ] Report specific cases that failed
- [ ] Adjust detection thresholds if necessary

---

## 🎯 RECOMMENDED WORKFLOW

```bash
# 1. Test footer removal safely
python remove_footer_v2.py
# Review output, look for problems

# 2. If OK, enable live mode
# Edit: DRY_RUN = False in remove_footer_v2.py
python remove_footer_v2.py

# 3. Extract activities
python extract_v2.py

# 4. Clean and validate
python cleanup_v2.py

# 5. Review validation report
cat validation_report.txt

# 6. Manually fix records needing review
# (Open output_clean.xlsx, filter needs_review = True)

# 7. Import to system
```

---

## 🔧 TUNING PARAMETERS

### If footer removal is too aggressive:
In `remove_footer_v2.py`:
- Increase `MAX_FOOTER_SCAN_ROWS` (default: 20)
- Tighten `looks_like_footer_row()` conditions

### If footer removal is too conservative:
- Decrease `MAX_FOOTER_SCAN_ROWS`
- Loosen `looks_like_footer_row()` conditions

### If don_vi detection misses units:
In `extract_v2.py`:
- Add keywords to `unit_keywords` list
- Lower `uppercase_ratio` threshold (currently 0.4)

### If too many validation failures:
In `cleanup_v2.py`:
- Adjust range limits (muc: 1-10, diem: 0-100)
- Adjust length limits

---

## 📈 PERFORMANCE COMPARISON

| Metric | v1 | v2 | Improvement |
|--------|----|----|-------------|
| Safety | ❌ No backups | ✅ Auto backup | Critical |
| Multi-activity | ❌ Excel only 1 | ✅ Properly handles | Major |
| Validation | ❌ None | ✅ Full validation | Major |
| Logging | ⚠️ Basic | ✅ Detailed stats | Good |
| Error handling | ⚠️ Fragile | ✅ Robust | Good |

---

## 🆘 TROUBLESHOOTING

### "Files are still corrupted"
- Did you run in DRY_RUN mode first?
- Check .backup files exist
- Restore from backups: `cp file.xlsx.backup file.xlsx`

### "Missing activities"
- Check statistics output
- Look for "Files with multiple activities"
- Review specific files manually

### "Too many validation failures"
- Check validation_report.txt
- May need to adjust thresholds
- Some records may legitimately need review

### "Student data was removed"
- THIS SHOULD NOT HAPPEN
- Check the removed rows in dry run output
- If it did happen, restore from .backup and report the case
