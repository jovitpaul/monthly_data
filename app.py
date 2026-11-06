import streamlit as st
import openpyxl
import difflib
import pandas as pd
import io

# --- HELPER FUNCTIONS ---

def get_similarity(str1, str2):
    s1, s2 = str1.lower().strip(), str2.lower().strip()
    seq_ratio = difflib.SequenceMatcher(None, s1, s2).ratio()
    
    words1 = sorted(s1.replace("-", " ").split())
    words2 = sorted(s2.replace("-", " ").split())
    sorted_ratio = difflib.SequenceMatcher(None, " ".join(words1), " ".join(words2)).ratio()
    
    return max(seq_ratio, sorted_ratio)

def extract_sheet_data(sheet):
    daily_volume = 0
    daily_manhours = 0
    vol_cell = "Not Found"
    man_cell = "Not Found"
    
    # 1. Total Volume
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

    # 2. Total Manhours
    found_manhours = False
    for row in range(1, sheet.max_row + 1):
        for col in range(1, sheet.max_column + 1):
            cell_val = sheet.cell(row=row, column=col).value
            if cell_val and str(cell_val).strip().lower() == "total manhours":
                for c in range(col + 1, col + 5):
                    adjacent_val = sheet.cell(row=row, column=c).value
                    if isinstance(adjacent_val, (int, float)):
                        daily_manhours = adjacent_val
                        col_letter = openpyxl.utils.get_column_letter(c)
                        man_cell = f"{col_letter}{row}"
                        found_manhours = True
                        break
            if found_manhours:
                break
        if found_manhours:
            break
            
    return daily_volume, daily_manhours, vol_cell, man_cell

def process_skus(sheet, monthly_totals):
    for row in range(1, sheet.max_row + 1):
        sku = sheet.cell(row=row, column=1).value
        qty = sheet.cell(row=row, column=2).value
        
        if not sku: continue
            
        sku_str = str(sku).strip()
        if sku_str.lower() in ["item description", "total", ""]: continue
            
        if isinstance(qty, (int, float)):
            if sku_str not in monthly_totals:
                monthly_totals[sku_str] = 0
            monthly_totals[sku_str] += qty

def merge_similar_skus(month1_totals, month2_totals, similarity_threshold=0.85):
    all_raw_skus = list(set(month1_totals.keys()).union(set(month2_totals.keys())))
    all_raw_skus.sort(key=len, reverse=True) 
    
    canonical_mapping = {}
    merged_data = {}
    
    for raw_sku in all_raw_skus:
        matched = False
        for canonical_sku in merged_data.keys():
            if get_similarity(raw_sku, canonical_sku) >= similarity_threshold:
                canonical_mapping[raw_sku] = canonical_sku
                matched = True
                break
                
        if not matched:
            canonical_mapping[raw_sku] = raw_sku
            merged_data[raw_sku] = {"Month 1": 0, "Month 2": 0}
            
    for sku, qty in month1_totals.items():
        canonical_name = canonical_mapping[sku]
        merged_data[canonical_name]["Month 1"] += qty
        
    for sku, qty in month2_totals.items():
        canonical_name = canonical_mapping[sku]
        merged_data[canonical_name]["Month 2"] += qty
        
    return merged_data

# --- DATA PROCESSING WRAPPERS ---

@st.cache_data(show_spinner=False)
def load_workbook_from_upload(uploaded_file):
    return openpyxl.load_workbook(uploaded_file, data_only=True)

def analyze_department_data(wb1, wb2):
    results = {"totals": {}, "skus": [], "audit1": [], "audit2": []}
    
    m1_vol_total, m1_man_total = 0, 0
    m2_vol_total, m2_man_total = 0, 0
    m1_skus, m2_skus = {}, {}
    
    # Process Month 1
    if wb1:
        for sheet_name in wb1.sheetnames:
            if sheet_name.strip().isdigit():
                sheet = wb1[sheet_name]
                v, m, v_cell, m_cell = extract_sheet_data(sheet)
                m1_vol_total += v
                m1_man_total += m
                results["audit1"].append({"Day": sheet_name, "Volume": v, "Vol Cell": v_cell, "Manhours": m, "Man Cell": m_cell})
                process_skus(sheet, m1_skus)
                
    # Process Month 2
    if wb2:
        for sheet_name in wb2.sheetnames:
            if sheet_name.strip().isdigit():
                sheet = wb2[sheet_name]
                v, m, v_cell, m_cell = extract_sheet_data(sheet)
                m2_vol_total += v
                m2_man_total += m
                results["audit2"].append({"Day": sheet_name, "Volume": v, "Vol Cell": v_cell, "Manhours": m, "Man Cell": m_cell})
                process_skus(sheet, m2_skus)
                
    # Compile Totals
    results["totals"] = {
        "Month 1 Volume": m1_vol_total, "Month 2 Volume": m2_vol_total, "Volume Diff": m2_vol_total - m1_vol_total,
        "Month 1 Manpower": m1_man_total, "Month 2 Manpower": m2_man_total, "Manpower Diff": m2_man_total - m1_man_total
    }
    
    # Compile SKUs
    merged_skus = merge_similar_skus(m1_skus, m2_skus, 0.85)
    for sku, data in merged_skus.items():
        if data["Month 1"] > 0 or data["Month 2"] > 0:
            results["skus"].append({
                "SKU (Cleaned)": sku,
                "Month 1 Total": data["Month 1"],
                "Month 2 Total": data["Month 2"],
                "Difference": data["Month 2"] - data["Month 1"]
            })
    results["skus"].sort(key=lambda x: x["Difference"], reverse=True)
    
    return results

# --- STREAMLIT UI ---

st.set_page_config(page_title="Monthly Data Analyzer", layout="wide")
st.title("📊 Monthly Data Analyzer")

# Sidebar for configuration and uploads
with st.sidebar:
    st.header("1. Settings")
    month1_name = st.text_input("Base Month Name (e.g., September)", value="September")
    month2_name = st.text_input("Comparison Month Name (e.g., October)", value="October")
    
    st.header(f"2. Upload Bakery Files")
    bakery_m1 = st.file_uploader(f"Bakery - {month1_name}", type=["xlsx"], key="bakery1")
    bakery_m2 = st.file_uploader(f"Bakery - {month2_name}", type=["xlsx"], key="bakery2")
    
    st.header(f"3. Upload Central Kitchen Files")
    ck_m1 = st.file_uploader(f"Central Kitchen - {month1_name}", type=["xlsx"], key="ck1")
    ck_m2 = st.file_uploader(f"Central Kitchen - {month2_name}", type=["xlsx"], key="ck2")

if st.button("Run Analysis", type="primary"):
    with st.spinner("Parsing Excel files..."):
        # Load files into memory
        wb_bakery_1 = load_workbook_from_upload(bakery_m1) if bakery_m1 else None
        wb_bakery_2 = load_workbook_from_upload(bakery_m2) if bakery_m2 else None
        wb_ck_1 = load_workbook_from_upload(ck_m1) if ck_m1 else None
        wb_ck_2 = load_workbook_from_upload(ck_m2) if ck_m2 else None
        
        # Process data
        bakery_data = analyze_department_data(wb_bakery_1, wb_bakery_2)
        ck_data = analyze_department_data(wb_ck_1, wb_ck_2)
        
        # UI TABS
        tab1, tab2, tab3 = st.tabs(["📈 Monthly Totals", "📦 SKU Comparison", "🔍 Data Validator (Audit)"])
        
        # --- TAB 1: MONTHLY TOTALS ---
        with tab1:
            col1, col2 = st.columns(2)
            with col1:
                st.subheader("🥐 Bakery")
                st.metric(label=f"Total Volume ({month2_name} vs {month1_name})", 
                          value=f"{bakery_data['totals']['Month 2 Volume']:,}", 
                          delta=f"{bakery_data['totals']['Volume Diff']:,}")
                st.metric(label=f"Total Manpower ({month2_name} vs {month1_name})", 
                          value=f"{bakery_data['totals']['Month 2 Manpower']:,}", 
                          delta=f"{bakery_data['totals']['Manpower Diff']:,}")
                          
            with col2:
                st.subheader("🍳 Central Kitchen")
                st.metric(label=f"Total Volume ({month2_name} vs {month1_name})", 
                          value=f"{ck_data['totals']['Month 2 Volume']:,}", 
                          delta=f"{ck_data['totals']['Volume Diff']:,}")
                st.metric(label=f"Total Manpower ({month2_name} vs {month1_name})", 
                          value=f"{ck_data['totals']['Month 2 Manpower']:,}", 
                          delta=f"{ck_data['totals']['Manpower Diff']:,}")

        # --- TAB 2: SKU COMPARISON ---
        with tab2:
            st.info("SKUs are merged if they are at least 85% similar. Sorted by highest positive difference.")
            col_b, col_c = st.columns(2)
            
            with col_b:
                st.subheader("Bakery SKUs")
                if bakery_data["skus"]:
                    df_bakery_skus = pd.DataFrame(bakery_data["skus"])
                    df_bakery_skus.rename(columns={"Month 1 Total": f"{month1_name} Total", "Month 2 Total": f"{month2_name} Total"}, inplace=True)
                    st.dataframe(df_bakery_skus, use_container_width=True)
                else:
                    st.write("No SKU data found.")
                    
            with col_c:
                st.subheader("Central Kitchen SKUs")
                if ck_data["skus"]:
                    df_ck_skus = pd.DataFrame(ck_data["skus"])
                    df_ck_skus.rename(columns={"Month 1 Total": f"{month1_name} Total", "Month 2 Total": f"{month2_name} Total"}, inplace=True)
                    st.dataframe(df_ck_skus, use_container_width=True)
                else:
                    st.write("No SKU data found.")

        # --- TAB 3: VALIDATOR ---
        with tab3:
            st.info("Check this log to see exactly which cells the script pulled the totals from on a day-by-day basis.")
            
            audit_tabs = st.tabs(["Bakery Base", "Bakery Comparison", "CK Base", "CK Comparison"])
            
            with audit_tabs[0]:
                st.write(f"**Bakery - {month1_name}**")
                st.dataframe(pd.DataFrame(bakery_data["audit1"]), use_container_width=True)
                
            with audit_tabs[1]:
                st.write(f"**Bakery - {month2_name}**")
                st.dataframe(pd.DataFrame(bakery_data["audit2"]), use_container_width=True)
                
            with audit_tabs[2]:
                st.write(f"**Central Kitchen - {month1_name}**")
                st.dataframe(pd.DataFrame(ck_data["audit1"]), use_container_width=True)
                
            with audit_tabs[3]:
                st.write(f"**Central Kitchen - {month2_name}**")
                st.dataframe(pd.DataFrame(ck_data["audit2"]), use_container_width=True)