import os
import re
import shutil
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from copy import copy
import warnings

# Suppress openpyxl warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, "input")
OUTPUT_DIR = os.path.join(BASE_DIR, "split_output")

KEY_PHRASE = "tên hoạt động:"


def sanitize_filename(name):
    """Remove invalid characters and limit filename length"""
    # Replace invalid characters with underscore
    name = re.sub(r'[<>:"/\\|?*]', "_", name)
    # Remove leading/trailing spaces and dots
    name = name.strip(". ")
    # Limit length
    if len(name) > 100:
        name = name[:100]
    return name


def extract_activity_name_from_sheet(ws):
    """Try to extract activity name from worksheet"""
    for row_idx in range(1, min(16, ws.max_row + 1)):  # Check first 15 rows
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row_idx, col_idx)
            cell_value = str(cell.value or "").strip().lower()
            
            if KEY_PHRASE in cell_value:
                # Try to get value from same cell after ":"
                if ":" in cell_value:
                    value = cell_value.split(":", 1)[1].strip()
                    if value and value != "none":
                        return value.title()
                
                # Try next column
                if col_idx + 1 <= ws.max_column:
                    next_cell = ws.cell(row_idx, col_idx + 1)
                    next_value = str(next_cell.value or "").strip()
                    if next_value and next_value.lower() != "none":
                        return next_value
                
                # Try same column, next row
                if row_idx + 1 <= ws.max_row:
                    next_row_cell = ws.cell(row_idx + 1, col_idx)
                    next_row_value = str(next_row_cell.value or "").strip()
                    if next_row_value and next_row_value.lower() != "none":
                        return next_row_value
    
    return None


def copy_sheet_with_formatting(source_sheet, target_workbook):
    """Copy sheet with all formatting to a new workbook"""
    target_sheet = target_workbook.active
    target_sheet.title = source_sheet.title
    
    # Copy column widths
    for col_letter in source_sheet.column_dimensions:
        if col_letter in source_sheet.column_dimensions:
            target_sheet.column_dimensions[col_letter].width = source_sheet.column_dimensions[col_letter].width
    
    # Copy row heights
    for row_num in source_sheet.row_dimensions:
        if row_num in source_sheet.row_dimensions:
            target_sheet.row_dimensions[row_num].height = source_sheet.row_dimensions[row_num].height
    
    # Copy cells with values and formatting (before merging)
    for row in source_sheet.iter_rows():
        for cell in row:
            target_cell = target_sheet[cell.coordinate]
            
            # Skip merged cells (they are read-only)
            if isinstance(cell, type(cell)) and hasattr(cell, 'value'):
                try:
                    target_cell.value = cell.value
                except AttributeError:
                    # Skip if it's a MergedCell
                    pass
            
            # Copy formatting
            if cell.has_style:
                try:
                    target_cell.font = copy(cell.font)
                    target_cell.border = copy(cell.border)
                    target_cell.fill = copy(cell.fill)
                    target_cell.number_format = copy(cell.number_format)
                    target_cell.protection = copy(cell.protection)
                    target_cell.alignment = copy(cell.alignment)
                except:
                    pass
    
    # Copy merged cells after all cells are copied
    for merged_cell_range in source_sheet.merged_cells.ranges:
        target_sheet.merge_cells(str(merged_cell_range))
    
    # Copy sheet properties
    try:
        target_sheet.sheet_format = copy(source_sheet.sheet_format)
        target_sheet.sheet_properties = copy(source_sheet.sheet_properties)
        target_sheet.page_setup = copy(source_sheet.page_setup)
        target_sheet.print_options = copy(source_sheet.print_options)
    except:
        pass
    
    return target_sheet


def split_excel_file(file_path, rel_path):
    """Split an Excel file into separate files per sheet"""
    try:
        # Load workbook with openpyxl to preserve formatting
        wb = load_workbook(file_path)
        
        if len(wb.sheetnames) < 2:
            print(f"⏭️  Skipping (only 1 sheet): {rel_path}")
            wb.close()
            return 0, []
        
        # Create output directory structure
        file_output_dir = os.path.join(OUTPUT_DIR, rel_path)
        os.makedirs(file_output_dir, exist_ok=True)
        
        split_count = 0
        split_files = []  # Track created files
        
        for sheet_name in wb.sheetnames:
            try:
                source_sheet = wb[sheet_name]
                
                # Try to extract activity name
                activity_name = extract_activity_name_from_sheet(source_sheet)
                
                if activity_name:
                    filename = sanitize_filename(activity_name)
                else:
                    filename = sanitize_filename(sheet_name)
                
                # Ensure unique filename
                output_path = os.path.join(file_output_dir, f"{filename}.xlsx")
                counter = 1
                while os.path.exists(output_path):
                    output_path = os.path.join(file_output_dir, f"{filename}_{counter}.xlsx")
                    counter += 1
                
                # Create new workbook and copy sheet with formatting
                new_wb = Workbook()
                new_wb.remove(new_wb.active)  # Remove default sheet
                new_wb.create_sheet(sheet_name)
                copy_sheet_with_formatting(source_sheet, new_wb)
                
                # Save new workbook
                new_wb.save(output_path)
                new_wb.close()
                
                file_name = os.path.basename(output_path)
                print(f"   ✅ Sheet '{sheet_name}' → {file_name}")
                split_files.append({"sheet": sheet_name, "file": file_name, "path": os.path.relpath(output_path, OUTPUT_DIR)})
                split_count += 1
                
            except Exception as e:
                print(f"   ❌ Error processing sheet '{sheet_name}': {e}")
        
        wb.close()
        return split_count, split_files
        
    except Exception as e:
        print(f"❌ Error opening file {rel_path}: {e}")
        return 0, []


def main():
    if not os.path.exists(INPUT_DIR):
        print(f"❌ Input directory not found: {INPUT_DIR}")
        return
    
    print(f"🔧 Processing files from: {INPUT_DIR}")
    print(f"📁 Output directory: {OUTPUT_DIR}\n")
    
    total_split = 0
    total_sheets = 0
    total_copied = 0
    split_summary = []  # Track all split operations
    
    for root, dirs, files in os.walk(INPUT_DIR):
        # Skip hidden and system directories
        dirs[:] = [d for d in dirs if not d.startswith(".") and d not in ["venv", "__pycache__"]]
        
        for file in files:
            file_path = os.path.join(root, file)
            rel_path = os.path.relpath(root, INPUT_DIR)
            rel_file = os.path.relpath(file_path, INPUT_DIR)
            
            # Prepare output directory
            file_output_dir = os.path.join(OUTPUT_DIR, rel_path)
            os.makedirs(file_output_dir, exist_ok=True)
            
            # Process Excel files
            if file.lower().endswith((".xlsx", ".xls")):
                print(f"📄 Processing Excel: {rel_file}")
                count, split_files = split_excel_file(file_path, rel_path)
                
                if count > 0:
                    # Multi-sheet file was split
                    total_split += 1
                    total_sheets += count
                    split_summary.append({
                        "source": rel_file,
                        "files": split_files
                    })
                else:
                    # Single-sheet file - copy as is
                    output_path = os.path.join(file_output_dir, file)
                    try:
                        shutil.copy2(file_path, output_path)
                        print(f"   📋 Copied (1 sheet): {file}")
                        total_copied += 1
                    except Exception as e:
                        print(f"   ❌ Error copying: {e}")
            else:
                # Non-Excel file - copy as is
                output_path = os.path.join(file_output_dir, file)
                try:
                    shutil.copy2(file_path, output_path)
                    print(f"📋 Copied: {rel_file}")
                    total_copied += 1
                except Exception as e:
                    print(f"❌ Error copying {rel_file}: {e}")
    
    print("\n" + "=" * 60)
    print(f"✅ DONE!")
    print(f"📊 Excel files split: {total_split} files → {total_sheets} sheets")
    print(f"📋 Files copied: {total_copied}")
    print(f"📁 Output saved to: {OUTPUT_DIR}")
    
    # Display split summary
    if split_summary:
        print("\n" + "=" * 60)
        print("📑 TỔNG KẾT CÁC FILE ĐÃ TÁCH:")
        print("=" * 60)
        
        # Group by folder
        folder_groups = {}
        for item in split_summary:
            folder = os.path.dirname(item['source'])
            if not folder:
                folder = "."
            if folder not in folder_groups:
                folder_groups[folder] = []
            folder_groups[folder].append(item)
        
        # Display tree structure
        folder_counter = 1
        for folder, items in sorted(folder_groups.items()):
            print(f"\n📁 [{folder_counter}] {folder}")
            
            for idx, item in enumerate(items):
                source_file = os.path.basename(item['source'])
                num_sheets = len(item['files'])
                is_last_file = (idx == len(items) - 1)
                file_prefix = "└─" if is_last_file else "├─"
                
                print(f"  {file_prefix} 📊 {source_file} ({num_sheets} sheets)")
                
                for file_idx, split_file in enumerate(item['files']):
                    is_last_split = (file_idx == len(item['files']) - 1)
                    
                    if is_last_file:
                        split_prefix = "      └─" if is_last_split else "      ├─"
                    else:
                        split_prefix = "  │   └─" if is_last_split else "  │   ├─"
                    
                    print(f"{split_prefix} 📄 {split_file['file']}")
            
            folder_counter += 1
        
        print("\n" + "=" * 60)
    
    print("=" * 60)


if __name__ == "__main__":
    main()
