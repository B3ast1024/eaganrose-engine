import os
import streamlit as st
import pdfplumber
import pandas as pd
import re
from datetime import datetime
import io
from dateutil.relativedelta import relativedelta
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# =====================================================================
# 1. AUTO-THEMING ENGINE
# =====================================================================
os.makedirs(".streamlit", exist_ok=True)
theme_config = """
[theme]
primaryColor = "#8E2D28"
backgroundColor = "#FFFFFF"
secondaryBackgroundColor = "#F0F2F6"
textColor = "#1E1E1E"
font = "sans serif"
"""
with open(".streamlit/config.toml", "w") as f:
    f.write(theme_config)

# =====================================================================
# 2. PLATFORM CONFIGURATION & VENDOR REGISTRY
# =====================================================================
VENDOR_REGISTRY = {
    "0010006727": {
        "name": "DAESANG",
        "skus": {
            "1570571": 0.03, "1793017": 0.03, "1908395": 0.05, 
            "1990882": 0.05, "2059134": 0.05
        }
    }
}

st.set_page_config(page_title="Eaganrose Engine", layout="wide")

# =====================================================================
# 3. FRONTEND UI & BRANDING
# =====================================================================
try:
    st.image("eagan rose logo.png", width=250)
except Exception:
    pass

st.title("Eaganrose Reconciliation Engine")
st.markdown("Automated PDF Ingestion & Variance Auditing for Costco Vendors")
st.markdown("---")

# =====================================================================
# 4. CORE EXTRACTION ENGINE
# =====================================================================
class InvalidCostcoPOError(Exception):
    pass

def extract_costco_pdf_data(pdf_file_obj):
    text = ""
    try:
        with pdfplumber.open(pdf_file_obj) as pdf:
            for page in pdf.pages:
                extracted = page.extract_text()
                if extracted:
                    text += extracted + "\n"
    except Exception as e:
        raise InvalidCostcoPOError(f"File is corrupted or not a readable PDF. ({str(e)})")
        
    order_match = re.search(r"Order\s*#[:\s\n]*(\d+)", text, re.IGNORECASE)
    if not order_match:
        raise InvalidCostcoPOError("Missing 'Order #'. This does not appear to be a valid Costco PO.")
        
    # --- HEADER EXTRACTION ---
    header_data = {}
    po_number = order_match.group(1)
    header_data['PO #'] = po_number
    header_data['Purchase Order #'] = po_number 
    
    detected_vendor_name = "UNKNOWN_VENDOR"
    detected_sku_map = {}
    for v_id, v_info in VENDOR_REGISTRY.items():
        if v_id in text:
            detected_vendor_name = v_info["name"]
            detected_sku_map = v_info["skus"]
            break
            
    header_data['Vendor'] = detected_vendor_name
    
    if len(po_number) >= 11:
        # Removed integer conversion to preserve leading zero
        header_data['Extracted PO Date'] = po_number[-7:-3] 
        header_data['PO Sequence'] = po_number[-3:]
        header_data['Extracted Depot'] = po_number[:-7].lstrip('0') 
    else:
        header_data['Extracted PO Date'], header_data['PO Sequence'], header_data['Extracted Depot'] = "", "", ""

    dept_match = re.search(r"Department\s*#[:\s\n]*([^\n]+)", text, re.IGNORECASE)
    if dept_match:
        parts = [p.strip() for p in dept_match.group(1).split('/')]
        raw_reg = parts[1] if len(parts) >= 2 else parts[0]
        reg_match = re.search(r"([A-Z]{2})", raw_reg)
        header_data['Region'] = reg_match.group(1) if reg_match else raw_reg[:2]
    else:
        header_data['Region'] = ""
    
    dates = re.findall(r"(\d{2}/\d{2}/\d{4})", text)
    if len(dates) >= 2:
        po_dt = datetime.strptime(dates[0], "%m/%d/%Y")
        del_dt = datetime.strptime(dates[1], "%m/%d/%Y")
        header_data['PO Date'] = po_dt.strftime("%Y-%m-%d")
        header_data[' PO Del Date'] = del_dt.strftime("%Y-%m-%d")
        header_data['Cancel Date'] = del_dt.strftime("%Y-%m-%d") 
        est_pay = (del_dt + relativedelta(months=1)).replace(day=1)
        header_data["Est'd  Month   Pay'd"] = est_pay.strftime("%Y-%m-%d")
    else:
        header_data['PO Date'], header_data[' PO Del Date'], header_data['Cancel Date'], header_data["Est'd  Month   Pay'd"] = "", "", "", ""
    
    # --- MULTI-ITEM EXTRACTION LOOP ---
    items_data = []
    
    item_matches = list(re.finditer(r"^(\d{1,3})\s+(\d{7})\s+.*?([A-Za-z].*?)\s+([\d\,\.]+)\s+(\d+)\s+CA", text, re.MULTILINE))
    
    if not item_matches:
        raise InvalidCostcoPOError("Could not isolate line items. The PDF formatting may be unexpected.")
        
    for i, match in enumerate(item_matches):
        item_dict = header_data.copy()
        
        sku = match.group(2)
        desc = match.group(3).strip().title()
        unit_cost = float(match.group(4).replace(',', ''))
        qty = int(match.group(5))
        
        item_dict['Costco Item #'] = sku
        item_dict['Item Description'] = desc
        item_dict['Unit Cost'] = unit_cost
        item_dict['Qty'] = qty
        item_dict['Com   %   Rate'] = detected_sku_map.get(sku, 0.0)
        item_dict['Spoils'] = 0.0075  
        
        # Isolate text chunk for allowances
        chunk_start = match.end()
        chunk_end = item_matches[i+1].start() if i + 1 < len(item_matches) else len(text)
        item_chunk = text[chunk_start:chunk_end]
        
        demo_total = 0.0
        allowance_lines = re.findall(r"^.*(?:ALLOWANCE|DFI\$).*$", item_chunk, re.IGNORECASE | re.MULTILINE)
        
        for line in allowance_lines:
            if re.search(r"SPOIL", line, re.IGNORECASE):
                continue
                
            amounts = re.findall(r"[\d\,]+\.\d{2}", line)
            if amounts:
                first_amount = float(amounts[0].replace(',', ''))
                last_amount = float(amounts[-1].replace(',', ''))
                
                if first_amount < 2.0:
                    continue
                    
                demo_total += last_amount
                    
        item_dict['Demo Accrual Deduction'] = demo_total
        
        freight_match = re.search(r"FREIGHT.*?(\d+\.\d{2})", item_chunk, re.IGNORECASE)
        freight_allowance = float(freight_match.group(1)) * qty if freight_match else 0.0
        item_dict['Freight Allowance'] = freight_allowance
        
        # Calculate True Net Inventory Amount (Gross - Allowances)
        gross_amt = unit_cost * qty
        spoils_calc = gross_amt * 0.0075
        item_dict['Net Inv Amt'] = gross_amt - demo_total - spoils_calc - freight_allowance
        
        items_data.append(item_dict)
        
    return items_data

# =====================================================================
# 5. DATA ROUTING & CALCULATIONS
# =====================================================================
def process_and_merge(df):
    bd_mask = df['Region'].str.contains('BD', case=False, na=False)
    bd_data = df[bd_mask].copy()
    d12_data = df[~bd_mask].copy()
    
    if not bd_data.empty:
        bd_data['Dead Net Cost'] = bd_data['Unit Cost'] - bd_data['Freight Allowance'] - bd_data['Demo Accrual Deduction']
        bd_data['Total Dead Net Amt'] = bd_data['Dead Net Cost'] * bd_data['Qty']
        
    master_df = pd.concat([bd_data, d12_data], ignore_index=True)
    if 'PO Date' in master_df.columns:
        master_df = master_df.sort_values(by='PO Date', ascending=True)
    return master_df

# =====================================================================
# 6. EXCEL EXPORT ENGINE
# =====================================================================
def parse_date(date_str):
    if not date_str: return None
    try: return datetime.strptime(date_str, "%Y-%m-%d").date()
    except: return date_str

def create_excel_buffer(master_df):
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Master_Ledger" 
    
    arial_font = Font(name='Arial', size=8)
    header_font = Font(name='Arial', size=8, bold=True, color="000000")
    header_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")
    right_align = Alignment(horizontal="right", vertical="center")
    
    headers = [
        'PO Date', 'Vendor', 'Region', 'Purchase Order #', '', '', '', 
        'Costco Item #', 'Item Description', 'Cancel Date', ' PO Del Date', 
        'Qty', 'Unit Cost', 'Spoils', 'Freight Allowance', 'Demo Accrual Deduction', 
        'Net Inv Amt', 'Com   %   Rate', 'Total Comision Earned', 
        "Est'd  Month   Pay'd", 'Date Comm  Paid', 'Ck #', 'Total Comision Paid', 
        'Commission Paid Difference from Commission Earned', 'Comments'
    ]
    
    ws.append(headers)
    for col_num, cell in enumerate(ws[1], 1):
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border
        
    for idx, row in master_df.iterrows():
        ws.append([
            parse_date(row.get('PO Date', '')), row.get('Vendor', ''), row.get('Region', ''),
            row.get('Extracted Depot', ''), row.get('Extracted PO Date', ''), row.get('PO Sequence', ''), '',
            row.get('Costco Item #', ''), row.get('Item Description', ''), parse_date(row.get('Cancel Date', '')),
            parse_date(row.get(' PO Del Date', '')), row.get('Qty', 0), row.get('Unit Cost', 0.0),
            row.get('Spoils', 0.0), row.get('Freight Allowance', 0.0), row.get('Demo Accrual Deduction', 0.0),
            row.get('Net Inv Amt', 0.0), row.get('Com   %   Rate', 0.0), 0.0,
            parse_date(row.get("Est'd  Month   Pay'd", '')), None, None, None, 0.0, ''
        ])
        
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=25):
        for cell in row:
            cell.font = arial_font
            cell.border = thin_border
            cell.alignment = center_align 
            
            col = cell.column
            
            if col in [4, 24]: 
                cell.alignment = right_align
            elif col in [6, 9]: 
                cell.alignment = left_align
            
            if col == 5: 
                cell.number_format = '0000' # Explicitly lock Column E to a 4-digit format
            elif col in [13, 23]: 
                cell.number_format = '0.00'
            elif col == 12: 
                cell.number_format = '#,##0' 
            elif col in [14, 18]: 
                cell.number_format = '0.00%'
            elif col in [15, 16]: 
                cell.number_format = '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
            elif col in [17, 19, 24]: 
                cell.number_format = '"$"#,##0.00'
            elif col in [1, 10, 11]: 
                cell.number_format = 'dd-mmm-yy'
            elif col == 20: 
                cell.number_format = 'mmm-yy'
                
        row_idx = row[0].row
        ws[f'S{row_idx}'] = f'=Q{row_idx}*R{row_idx}'
        ws[f'X{row_idx}'] = f'=W{row_idx}-S{row_idx}'

    wb.save(output)
    output.seek(0)
    return output

# =====================================================================
# 7. APP EXECUTION TRIGGER
# =====================================================================
uploaded_files = st.file_uploader("Upload Costco POs (PDF)", type="pdf", accept_multiple_files=True)

if uploaded_files:
    if st.button("Run Audit Engine"):
        with st.spinner(f"Ingesting {len(uploaded_files)} Purchase Orders..."):
            all_data, failed_files = [], []
            for pdf_file in uploaded_files:
                try:
                    extracted_items = extract_costco_pdf_data(pdf_file)
                    all_data.extend(extracted_items)
                except Exception as e:
                    failed_files.append({"filename": pdf_file.name, "error": str(e)})
            
            if all_data:
                master_df = process_and_merge(pd.DataFrame(all_data))
                st.success(f"Audit Complete! Processed {len(master_df)} valid records.")
                st.download_button(
                    label="📥 Download Master Ledger (Excel)",
                    data=create_excel_buffer(master_df),
                    file_name="Audit Software: Delete After Use.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            if failed_files:
                st.warning(f"⚠️ Skipped {len(failed_files)} file(s).")
                with st.expander("View Error Log"):
                    for fail in failed_files: st.write(f"- **{fail['filename']}**: {fail['error']}")