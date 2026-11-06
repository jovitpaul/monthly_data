import openpyxl
from pathlib import Path
import difflib

# Define the path to your folder
data_folder = Path(r"C:\Users\MANUFACTURING\Desktop\monthly_data")

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

def get_similarity(str1, str2):
    """
    Calculates similarity between two strings. 
    Checks both exact sequence and word-sorted sequence to handle swapped words.
    """
    s1, s2 = str1.lower().strip(), str2.lower().strip()
    
    # 1. Straight sequence similarity
    seq_ratio = difflib.SequenceMatcher(None, s1, s2).ratio()
    
    # 2. Sorted word similarity (Handles "Pandan Macaroons" vs "Macaroons - Pandan")
    words1 = sorted(s1.replace("-", " ").split())
    words2 = sorted(s2.replace("-", " ").split())
    sorted_s1 = " ".join(words1)
    sorted_s2 = " ".join(words2)
    sorted_ratio = difflib.SequenceMatcher(None, sorted_s1, sorted_s2).ratio()
    
    return max(seq_ratio, sorted_ratio)

def extract_sku_data(sheet, monthly_totals):
    """
    Parses a single sheet to accumulate raw SKU quantities.
    """
    for row in range(1, sheet.max_row + 1):
        sku = sheet.cell(row=row, column=1).value
        qty = sheet.cell(row=row, column=2).value
        
        if not sku:
            continue
            
        sku_str = str(sku).strip()
        if sku_str.lower() in ["item description", "total", ""]:
            continue
            
        if isinstance(qty, (int, float)):
            if sku_str not in monthly_totals:
                monthly_totals[sku_str] = 0
            monthly_totals[sku_str] += qty

def process_month_skus(filepath):
    monthly_totals = {}
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
    except FileNotFoundError:
        print(f"Warning: Could not find file {filepath}")
        return monthly_totals

    for sheet_name in wb.sheetnames:
        if sheet_name.strip().isdigit():
            sheet = wb[sheet_name]
            extract_sku_data(sheet, monthly_totals)
            
    return monthly_totals

def merge_similar_skus(aug_totals, sep_totals, similarity_threshold=0.85):
    """
    Groups SKUs from both months that are >= 85% similar into single canonical names.
    """
    all_raw_skus = list(set(aug_totals.keys()).union(set(sep_totals.keys())))
    
    # Sort by length descending so the longest, most descriptive name becomes the "Canonical" name
    all_raw_skus.sort(key=len, reverse=True) 
    
    canonical_mapping = {}
    merged_data = {}
    
    for raw_sku in all_raw_skus:
        matched = False
        
        # Check against existing established canonical SKUs
        for canonical_sku in merged_data.keys():
            if get_similarity(raw_sku, canonical_sku) >= similarity_threshold:
                canonical_mapping[raw_sku] = canonical_sku
                matched = True
                break
                
        # If no 85% match is found, establish it as a new canonical SKU
        if not matched:
            canonical_mapping[raw_sku] = raw_sku
            merged_data[raw_sku] = {"August": 0, "September": 0}
            
    # Apply mapped totals
    for sku, qty in aug_totals.items():
        canonical_name = canonical_mapping[sku]
        merged_data[canonical_name]["August"] += qty
        
    for sku, qty in sep_totals.items():
        canonical_name = canonical_mapping[sku]
        merged_data[canonical_name]["September"] += qty
        
    return merged_data

def main():
    print("--- Generating Cleaned SKU Comparison Report (85% Similarity Merge) ---")
    
    wb_out = openpyxl.Workbook()
    wb_out.remove(wb_out.active)
    
    for department, months in files.items():
        print(f"\nProcessing SKUs for {department}...")
        
        aug_path = data_folder / months["August"]
        aug_totals = process_month_skus(aug_path)
        
        sep_path = data_folder / months["September"]
        sep_totals = process_month_skus(sep_path)
        
        # Merge SKUs using 85% similarity rule
        merged_skus = merge_similar_skus(aug_totals, sep_totals, 0.85)
        
        comparison_list = []
        for sku, data in merged_skus.items():
            aug_qty = data["August"]
            sep_qty = data["September"]
            diff = sep_qty - aug_qty
            
            # Only include SKUs that had activity in at least one of the months
            if aug_qty > 0 or sep_qty > 0:
                comparison_list.append({
                    "SKU": sku,
                    "August_Total": aug_qty,
                    "September_Total": sep_qty,
                    "Difference": diff
                })
            
        # Sort from highest positive difference to lowest
        comparison_list.sort(key=lambda x: x["Difference"], reverse=True)
        
        ws_out = wb_out.create_sheet(title=f"{department} Cleaned SKUs"[:31])
        
        headers = ["Item Description (Cleaned)", "August Total", "September Total", "Difference (Sep - Aug)"]
        ws_out.append(headers)
        
        for cell in ws_out[1]:
            cell.font = openpyxl.styles.Font(bold=True)
        
        for item in comparison_list:
            ws_out.append([item["SKU"], item["August_Total"], item["September_Total"], item["Difference"]])
            
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
            
        print(f"  -> Consolidated to {len(comparison_list)} unique SKUs.")

    output_file = data_folder / "Cleaned_SKU_Difference_Report.xlsx"
    wb_out.save(output_file)
    print(f"\nDone! Report successfully saved to: {output_file}")

if __name__ == "__main__":
    main()