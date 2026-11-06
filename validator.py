import openpyxl
from pathlib import Path

# Define the path to your folder
data_folder = Path(r"C:\Users\MANUFACTURING\Desktop\monthly_data")

files = {
    "August": "August_Bakery_2026.xlsx",
    "September": "September_Bakery_2026.xlsx"
}

def extract_and_audit(sheet, sheet_name):
    """
    Parses a single sheet and logs the exact cell coordinates it used.
    """
    daily_volume = 0
    daily_manhours = 0
    vol_cell = "Not Found"
    man_cell = "Not Found"
    
    # 1. Parse Total Volume
    for row in range(1, sheet.max_row + 1):
        cell_a = sheet.cell(row=row, column=1).value
        
        if cell_a and str(cell_a).strip().lower() == "total":
            for offset in range(0, 6):
                val_b = sheet.cell(row=row+offset, column=2).value
                if isinstance(val_b, (int, float)):
                    daily_volume = val_b
                    vol_cell = f"B{row+offset}"
                    break
            break 

    # 2. Parse Total Manhours
    found_manhours = False
    for row in range(1, sheet.max_row + 1):
        for col in range(1, sheet.max_column + 1):
            cell_val = sheet.cell(row=row, column=col).value
            
            if cell_val and str(cell_val).strip().lower() == "total manhours":
                for c in range(col + 1, col + 5):
                    adjacent_val = sheet.cell(row=row, column=c).value
                    if isinstance(adjacent_val, (int, float)):
                        daily_manhours = adjacent_val
                        # Convert column number to letter for easy reading (e.g., 7 -> G)
                        col_letter = openpyxl.utils.get_column_letter(c)
                        man_cell = f"{col_letter}{row}"
                        found_manhours = True
                        break
            if found_manhours:
                break
        if found_manhours:
            break
            
    print(f"  Day {sheet_name:>2}: Volume = {daily_volume:<8} (Cell {vol_cell:<4}) | Manhours = {daily_manhours:<6} (Cell {man_cell})")
    return daily_volume, daily_manhours

def process_audit_workbook(filepath, month_name):
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
    except FileNotFoundError:
        print(f"Warning: Could not find file {filepath}")
        return 0, 0

    monthly_volume = 0
    monthly_manhours = 0
    
    print(f"\n--- {month_name.upper()} BAKERY AUDIT ---")
    for sheet_name in wb.sheetnames:
        if sheet_name.strip().isdigit():
            vol, manhours = extract_and_audit(wb[sheet_name], sheet_name)
            monthly_volume += vol
            monthly_manhours += manhours
            
    print(f"  --> {month_name} TOTALS: Volume = {monthly_volume}, Manhours = {monthly_manhours}")
    return monthly_volume, monthly_manhours

def main():
    print("Starting Bakery Audit...\n")
    
    aug_path = data_folder / files["August"]
    process_audit_workbook(aug_path, "August")
    
    sep_path = data_folder / files["September"]
    process_audit_workbook(sep_path, "September")

if __name__ == "__main__":
    main()