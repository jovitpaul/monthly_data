import os
import openpyxl
from pathlib import Path

# Define the path to your folder
data_folder = Path(r"C:\Users\MANUFACTURING\Desktop\monthly_data")

# Define the files based on your folder structure
files = {
    "Bakery": {
        "August": "August_Bakery_2026.xlsx",
        "September": "September_Bakery_2026.xlsx"
    },
    "Central Kitchen": {
        "August": "August_2026_Central_Kitchen.xlsx",
        "September": "September_2026_Central_Kitchen.xlsx"
    }
}

def extract_sheet_data(sheet):
    """
    Parses a single sheet to find Total Volume and Total Manhours.
    """
    daily_volume = 0
    daily_manhours = 0
    
    # 1. Parse Total Volume (Look for "Total" in Column A, value in Column B)
    for row in range(1, sheet.max_row + 1):
        cell_a = sheet.cell(row=row, column=1).value
        
        if cell_a and str(cell_a).strip().lower() == "total":
            # Scan down up to 5 rows in Col B to handle misaligned numbers
            for offset in range(0, 6):
                val_b = sheet.cell(row=row+offset, column=2).value
                if isinstance(val_b, (int, float)):
                    daily_volume = val_b
                    break
            break 

    # 2. Parse Total Manhours (Handle merged cells by scanning right)
    found_manhours = False
    for row in range(1, sheet.max_row + 1):
        for col in range(1, sheet.max_column + 1):
            cell_val = sheet.cell(row=row, column=col).value
            
            if cell_val and str(cell_val).strip().lower() == "total manhours":
                # Scan up to 4 columns to the right (e.g., from E to G) for the value
                for c in range(col + 1, col + 5):
                    adjacent_val = sheet.cell(row=row, column=c).value
                    if isinstance(adjacent_val, (int, float)):
                        daily_manhours = adjacent_val
                        found_manhours = True
                        break
            if found_manhours:
                break
        if found_manhours:
            break
            
    return daily_volume, daily_manhours

def process_workbook(filepath):
    """
    Iterates through daily sheets (1 to 31) and returns grand totals.
    """
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
    except FileNotFoundError:
        print(f"Warning: Could not find file {filepath}")
        return 0, 0

    monthly_volume = 0
    monthly_manhours = 0
    
    for sheet_name in wb.sheetnames:
        if sheet_name.strip().isdigit():
            sheet = wb[sheet_name]
            vol, manhours = extract_sheet_data(sheet)
            monthly_volume += vol
            monthly_manhours += manhours
            
    return monthly_volume, monthly_manhours

def main():
    print("--- Extracting and Calculating Data ---")
    
    results = []
    
    for department, months in files.items():
        print(f"Processing {department}...")
        
        # Calculate August
        aug_path = data_folder / months["August"]
        aug_vol, aug_man = process_workbook(aug_path)
        
        # Calculate September
        sep_path = data_folder / months["September"]
        sep_vol, sep_man = process_workbook(sep_path)
        
        # Differences
        vol_diff = sep_vol - aug_vol
        man_diff = sep_man - aug_man
        
        # Store for Excel
        results.append([department, "Total Volume", aug_vol, sep_vol, vol_diff])
        results.append([department, "Total Manhours", aug_man, sep_man, man_diff])

    # --- SAVE TO EXCEL ---
    output_file = data_folder / "Comparison_Report.xlsx"
    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active
    ws_out.title = "Monthly Comparison"
    
    # Write Headers
    headers = ["Department", "Metric", "August", "September", "Difference (Sep vs Aug)"]
    ws_out.append(headers)
    
    # Write Data
    for row_data in results:
        ws_out.append(row_data)
        
    # Auto-adjust column width for readability
    for col in ws_out.columns:
        max_length = 0
        column_letter = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        ws_out.column_dimensions[column_letter].width = max_length + 2

    wb_out.save(output_file)
    print(f"\nDone! Report successfully saved to: {output_file}")

if __name__ == "__main__":
    main()