import pandas as pd
import numpy as np
import re
import pdfplumber



def get_rounded_rate(tax_amt, taxable_val):
    if taxable_val == 0 or tax_amt == 0:
        return ''
    raw_rate = (tax_amt / taxable_val) * 100
    # Standard GST rates (both half rates for CGST/SGST and full for IGST)
    valid_rates = [2.5, 5, 6, 9, 12, 14, 18, 20, 28, 40]
    
    # Find the closest valid rate
    closest_rate = min(valid_rates, key=lambda x: abs(x - raw_rate))
    
    # Return as integer if it's a whole number (e.g. 9 instead of 9.0)
    if closest_rate == int(closest_rate):
        return int(closest_rate)
    return closest_rate

def get_col_index(df, possible_names, search_rows=10):
    for idx in range(len(df.columns)):
        for row_idx in range(min(search_rows, len(df))):
            val = str(df.iloc[row_idx, idx]).strip().lower()
            for name in possible_names:
                if name.lower() == val or name.lower() in val:
                    return idx
    return -1

def format_date(date_val):
    if pd.isna(date_val) or date_val == '':
        return ''
    try:
        if isinstance(date_val, pd.Timestamp):
            return date_val.strftime('%d-%m-%Y')
        else:
            return pd.to_datetime(date_val, dayfirst=True).strftime('%d-%m-%Y')
    except:
        return str(date_val)

def find_header_row(file_path, sheet_name, expected_col):
    df_raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None, nrows=10)
    for i, row in df_raw.iterrows():
        if any(expected_col.lower() == str(val).lower().strip() for val in row.values if pd.notna(val)):
            return i
    return 0

def clean_company_name(name):
    name = str(name).upper()
    name = re.sub(r'\b(PRIVATE|PVT|LIMITED|LTD)\b', '', name)
    return re.sub(r'[^A-Z0-9]', '', name)

def extract_tds_from_26as(pdf_path):
    if not pdf_path: return {}
    print(f"Extracting TDS deterministically from {pdf_path}...")
    tds_dict = {}
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        clean_row = [str(x).strip().replace('\n', ' ') for x in row if x is not None and str(x).strip() != '']
                        if len(clean_row) >= 5:
                            tan_idx = -1
                            for i, val in enumerate(clean_row):
                                if re.match(r'^[A-Z]{4}\d{5}[A-Z]$', val):
                                    tan_idx = i
                                    break
                            
                            if tan_idx != -1 and tan_idx >= 1:
                                name = clean_row[1] if tan_idx == 2 else " ".join(clean_row[1:tan_idx])
                                try:
                                    total_amount = float(clean_row[tan_idx + 1].replace(',', ''))
                                    total_tds = float(clean_row[tan_idx + 2].replace(',', ''))
                                    if total_amount > 0:
                                        rate = round((total_tds / total_amount) * 100)
                                        c_name = clean_company_name(name)
                                        tds_dict[c_name] = {'amount': total_amount, 'rate': rate, 'raw_name': name}
                                except Exception:
                                    pass
        print("Deterministic TDS extraction complete.")
    except Exception as e:
        print(f"Error extracting TDS from 26AS deterministically: {e}")
    return tds_dict

def process_sales_data(b2b_path, b2c_path, form_26as_path=None):
    """
    Reads GSTR1 B2B and B2C Excel files and maps them to the Tally Dashboard format.
    """
    final_rows = []
    
    tds_data = extract_tds_from_26as(form_26as_path)
    tds_balances = {k: v['amount'] for k, v in tds_data.items()}
    tds_rates = {k: v['rate'] for k, v in tds_data.items()}
    
    # --- 1. Process B2B Data ---
    try:
        xl = pd.ExcelFile(b2b_path)
        sheet_names = xl.sheet_names
        b2b_sheet = next((s for s in sheet_names if any(x in s.lower() for x in ['b2b', 'invoice', 'invoices', 'sheet1', 'sheet 1'])), sheet_names[0] if len(sheet_names) == 1 else None)
        
        if b2b_sheet:
            df_b2b_raw = pd.read_excel(b2b_path, sheet_name=b2b_sheet, header=None)
            h_idx_b2b = find_header_row(b2b_path, b2b_sheet, 'Invoice No')
            if h_idx_b2b == 0:
                h_idx_b2b = find_header_row(b2b_path, b2b_sheet, 'Recipients Name')

            party_col = get_col_index(df_b2b_raw, ['Receiver Name', 'Recipients Name', 'Party Name'])
            inv_num_col = get_col_index(df_b2b_raw, ['Invoice Number', 'Invoice No'])
            inv_date_col = get_col_index(df_b2b_raw, ['Invoice Date'])
            taxable_col = get_col_index(df_b2b_raw, ['Taxable Value', 'Taxable value'])
            cgst_col = get_col_index(df_b2b_raw, ['Central Tax', 'CGST', 'Central Tax Amount'])
            sgst_col = get_col_index(df_b2b_raw, ['State/UT Tax', 'SGST', 'State Tax Amount'])
            igst_col = get_col_index(df_b2b_raw, ['Integrated Tax', 'IGST', 'Integrated Tax Amount'])
            
            df_b2b = df_b2b_raw.iloc[h_idx_b2b+1:].dropna(subset=[inv_num_col]) if inv_num_col != -1 else df_b2b_raw.iloc[h_idx_b2b+1:]
            
            for _, row in df_b2b.iterrows():
                if inv_num_col != -1 and pd.isna(row[inv_num_col]): continue
                
                party_name = str(row[party_col]).strip() if party_col != -1 and pd.notna(row[party_col]) else ''
                try: taxable_val = float(row[taxable_col]) if taxable_col != -1 and pd.notna(row[taxable_col]) else 0
                except: taxable_val = 0
                try: cgst_amt = float(row[cgst_col]) if cgst_col != -1 and pd.notna(row[cgst_col]) else 0
                except: cgst_amt = 0
                try: sgst_amt = float(row[sgst_col]) if sgst_col != -1 and pd.notna(row[sgst_col]) else 0
                except: sgst_amt = 0
                try: igst_amt = float(row[igst_col]) if igst_col != -1 and pd.notna(row[igst_col]) else 0
                except: igst_amt = 0
                
                inv_num = str(row[inv_num_col]).strip() if inv_num_col != -1 and pd.notna(row[inv_num_col]) else ''
                inv_date = str(row[inv_date_col]).strip() if inv_date_col != -1 and pd.notna(row[inv_date_col]) else ''
                
                sales_ledger = "Sales"
                narration = f"Being sales made to {party_name}"
                
                tds_ledger = ''
                tds_rate_val = ''
                tds_note = ''
                c_party = clean_company_name(party_name)
                
                matched_tds_key = None
                for k in tds_rates:
                    if k in c_party or c_party in k:
                        matched_tds_key = k
                        break
                
                if matched_tds_key:
                    tds_ledger = 'TDS Receivable'
                    tds_rate_val = tds_rates[matched_tds_key]
                    tds_balances[matched_tds_key] -= taxable_val
                    if tds_balances[matched_tds_key] < 0:
                        tds_note = 'Not in 26AS yet'

                final_rows.append({
                    'Date': format_date(inv_date),
                    'Voucher Type': 'Sales',
                    'Voucher Number': inv_num,
                    'Party Ledger': party_name,
                    'Sales Ledger': sales_ledger,
                    'Sales Item Name': '',
                    'Sales Item Qty': '',
                    'Sales Item UOM': '',
                    'Sales Item Rate': '',
                    'Taxable Value': taxable_val,
                    'CGST Ledger': 'Output CGST' if cgst_amt > 0 else '',
                    'CGST Rate': get_rounded_rate(cgst_amt, taxable_val),
                    'SGST Ledger': 'Output SGST' if sgst_amt > 0 else '',
                    'SGST Rate': get_rounded_rate(sgst_amt, taxable_val),
                    'IGST Ledger': 'Output IGST' if igst_amt > 0 else '',
                    'IGST Rate': get_rounded_rate(igst_amt, taxable_val),
                    'Narration': narration,
                    'TDS Receivable Ledger': tds_ledger,
                    'TDS %': tds_rate_val,
                    'TDS Note': tds_note
                })
    except Exception as e:
        raise Exception(f"Error processing B2B Sales Data: {e}")

    # --- 1.5 Process B2B Notes (Credit Notes) ---
    try:
        xl = pd.ExcelFile(b2b_path)
        sheet_names = xl.sheet_names
        note_sheet = next((s for s in sheet_names if any(x in s.lower() for x in ['cdnr', 'note', 'credit note'])), None)
        
        if note_sheet:
            df_note_raw = pd.read_excel(b2b_path, sheet_name=note_sheet, header=None)
            h_idx_note = find_header_row(b2b_path, note_sheet, 'Note No')
            if h_idx_note == 0:
                h_idx_note = find_header_row(b2b_path, note_sheet, 'Recipients Name')

            party_col = get_col_index(df_note_raw, ['Receiver Name', 'Recipients Name', 'Party Name'])
            note_num_col = get_col_index(df_note_raw, ['Note Number', 'Note No', 'Credit Note No'])
            note_date_col = get_col_index(df_note_raw, ['Note Date', 'Credit Note Date'])
            taxable_col = get_col_index(df_note_raw, ['Taxable Value', 'Taxable value'])
            cgst_col = get_col_index(df_note_raw, ['Central Tax', 'CGST', 'Central Tax Amount'])
            sgst_col = get_col_index(df_note_raw, ['State/UT Tax', 'SGST', 'State Tax Amount'])
            igst_col = get_col_index(df_note_raw, ['Integrated Tax', 'IGST', 'Integrated Tax Amount'])
            
            df_note = df_note_raw.iloc[h_idx_note+1:].dropna(subset=[note_num_col]) if note_num_col != -1 else df_note_raw.iloc[h_idx_note+1:]
            
            for _, row in df_note.iterrows():
                if note_num_col != -1 and pd.isna(row[note_num_col]): continue
                
                party_name = str(row[party_col]).strip() if party_col != -1 and pd.notna(row[party_col]) else ''
                try: taxable_val = float(row[taxable_col]) if taxable_col != -1 and pd.notna(row[taxable_col]) else 0
                except: taxable_val = 0
                try: cgst_amt = float(row[cgst_col]) if cgst_col != -1 and pd.notna(row[cgst_col]) else 0
                except: cgst_amt = 0
                try: sgst_amt = float(row[sgst_col]) if sgst_col != -1 and pd.notna(row[sgst_col]) else 0
                except: sgst_amt = 0
                try: igst_amt = float(row[igst_col]) if igst_col != -1 and pd.notna(row[igst_col]) else 0
                except: igst_amt = 0
                
                note_num = str(row[note_num_col]).strip() if note_num_col != -1 and pd.notna(row[note_num_col]) else ''
                note_date = str(row[note_date_col]).strip() if note_date_col != -1 and pd.notna(row[note_date_col]) else ''
                
                sales_ledger = "Sales"

                final_rows.append({
                    'Date': format_date(note_date),
                    'Voucher Type': 'Credit Note',
                    'Voucher Number': note_num,
                    'Party Ledger': party_name,
                    'Sales Ledger': sales_ledger,
                    'Sales Item Name': '',
                    'Sales Item Qty': '',
                    'Sales Item UOM': '',
                    'Sales Item Rate': '',
                    'Taxable Value': taxable_val,
                    'CGST Ledger': 'Output CGST' if cgst_amt > 0 else '',
                    'CGST Rate': get_rounded_rate(cgst_amt, taxable_val),
                    'SGST Ledger': 'Output SGST' if sgst_amt > 0 else '',
                    'SGST Rate': get_rounded_rate(sgst_amt, taxable_val),
                    'IGST Ledger': 'Output IGST' if igst_amt > 0 else '',
                    'IGST Rate': get_rounded_rate(igst_amt, taxable_val),
                    'Narration': f"Being sales return / credit note for {party_name}",
                    'TDS Receivable Ledger': '',
                    'TDS %': '',
                    'TDS Note': ''
                })
    except Exception as e:
        raise Exception(f"Error processing B2B Credit Notes: {e}")

    # --- 2. Process B2C Data ---
    try:
        if b2c_path:
            df_b2c = pd.read_excel(b2c_path, header=None)
            
            # Determine the date for B2C based on the header info
            b2c_date = "01-01-2026" # Default fallback
            try:
                fy_text = str(df_b2c.iloc[0, 0]) # "INFINITIMINDS TECHNOLOGIES  (F.Y.:2025-2026)"
                month_text = str(df_b2c.iloc[1, 0]) # "B2c Small Mismatch (June)"
                
                month_match = re.search(r'\(([A-Za-z]+)\)', month_text)
                fy_match = re.search(r'F\.Y\.:(\d{4})-(\d{4})', fy_text)
                
                if month_match and fy_match:
                    month_str = month_match.group(1)
                    start_year = int(fy_match.group(1))
                    end_year = int(fy_match.group(2))
                    
                    # Convert month string to month number
                    month_num = pd.to_datetime(month_str, format='%B').month
                    
                    # If Jan, Feb, Mar, it's the end year of FY. Else start year.
                    year = end_year if month_num in [1, 2, 3] else start_year
                    b2c_date = f"01-{month_num:02d}-{year}"
            except Exception as e:
                print(f"Could not parse B2C date from headers: {e}")

            taxable_col = get_col_index(df_b2c, ['Taxable Value', 'Taxable value'])
            cgst_col = get_col_index(df_b2c, ['Central Tax', 'CGST', 'Central Tax Amount'])
            sgst_col = get_col_index(df_b2c, ['State/UT Tax', 'SGST', 'State Tax Amount'])
            igst_col = get_col_index(df_b2c, ['Integrated Tax', 'IGST', 'Integrated Tax Amount'])
            place_of_supply_col = get_col_index(df_b2c, ['Place of Supply', 'State'])
            
            # Use dynamic header row find for B2C data rows
            h_idx_b2c = find_header_row(b2c_path, 0, 'Taxable Value')
            data_rows = df_b2c.iloc[h_idx_b2c+1:].copy()
            
            b2c_counter = 1
            for _, row in data_rows.iterrows():
                if pd.isna(row.iloc[2]) or str(row.iloc[2]).strip().lower() == 'total' or (taxable_col != -1 and str(row.iloc[taxable_col]).strip() == 'Taxable Value'):
                    continue # Skip empty, total, or header rows
                    
                try: taxable_val = float(row.iloc[taxable_col]) if taxable_col != -1 and pd.notna(row.iloc[taxable_col]) else 0
                except: taxable_val = 0
                try: cgst_amt = float(row.iloc[cgst_col]) if cgst_col != -1 and pd.notna(row.iloc[cgst_col]) else 0
                except: cgst_amt = 0
                try: sgst_amt = float(row.iloc[sgst_col]) if sgst_col != -1 and pd.notna(row.iloc[sgst_col]) else 0
                except: sgst_amt = 0
                try: igst_amt = float(row.iloc[igst_col]) if igst_col != -1 and pd.notna(row.iloc[igst_col]) else 0
                except: igst_amt = 0
                
                if taxable_val == 0:
                    continue
                
                place_of_supply = str(row.iloc[place_of_supply_col]) if place_of_supply_col != -1 else str(row.iloc[2])
                sales_ledger = "Sales"

                final_rows.append({
                    'Date': b2c_date,
                    'Voucher Type': 'Sales',
                    'Voucher Number': str(b2c_counter),
                    'Party Ledger': 'B2C Debtors',
                    'Sales Ledger': sales_ledger,
                    'Sales Item Name': '',
                    'Sales Item Qty': '',
                    'Sales Item UOM': '',
                    'Sales Item Rate': '',
                    'Taxable Value': taxable_val,
                    'CGST Ledger': 'Output CGST' if cgst_amt > 0 else '',
                    'CGST Rate': get_rounded_rate(cgst_amt, taxable_val),
                    'SGST Ledger': 'Output SGST' if sgst_amt > 0 else '',
                    'SGST Rate': get_rounded_rate(sgst_amt, taxable_val),
                    'IGST Ledger': 'Output IGST' if igst_amt > 0 else '',
                    'IGST Rate': get_rounded_rate(igst_amt, taxable_val),
                    'Narration': f"Being B2C sales summary for {place_of_supply}",
                    'TDS Receivable Ledger': '',
                    'TDS %': '',
                    'TDS Note': ''
                })
                b2c_counter += 1
    except Exception as e:
        raise Exception(f"Error processing B2C Sales Data: {e}")

    # Convert to DataFrame
    df_final = pd.DataFrame(final_rows)
    
    # Extract unique ledgers and map them
    ledger_map = {}
    
    for row in final_rows:
        party = row.get('Party Ledger', '')
        if party: ledger_map[party] = 'Sundry Debtor'
        
        sales = row.get('Sales Ledger', '')
        if sales: ledger_map[sales] = 'Sales Accounts.'
        
        cgst = row.get('CGST Ledger', '')
        if cgst: ledger_map[cgst] = 'Duties & Taxes'
        
        sgst = row.get('SGST Ledger', '')
        if sgst: ledger_map[sgst] = 'Duties & Taxes'
        
        igst = row.get('IGST Ledger', '')
        if igst: ledger_map[igst] = 'Duties & Taxes'
        
        tds = row.get('TDS Receivable Ledger', '')
        if tds: ledger_map[tds] = 'Duties & Taxes'
        
    df_ledgers = pd.DataFrame([
        {'Ledger Name': k, 'Under Group': v} 
        for k, v in ledger_map.items()
    ])
    
    return df_final, df_ledgers

if __name__ == "__main__":
    b2b_file = r"C:\Users\Admin\.gemini\antigravity\scratch\tally_automation\Input\April May GSTR1.xlsx"
    b2c_file = None
    form_26as_file = r"C:\Users\Admin\.gemini\antigravity\scratch\tally_automation\Input\form 26 AS.pdf"
    
    result_df, ledgers_df = process_sales_data(b2b_file, b2c_file, form_26as_file)
    print("--- Final Processed Data (First 5 rows) ---")
    print(result_df.head().to_string())
    
    output_path = r"C:\Users\Admin\.gemini\antigravity\scratch\tally_automation\Output\processed_sales.xlsx"
    with pd.ExcelWriter(output_path) as writer:
        result_df.to_excel(writer, sheet_name='RAW_DATA_MASTER', index=False)
        ledgers_df.to_excel(writer, sheet_name='LEDGER_GROUP_MAP', index=False)
    print(f"\nSaved updated test output to {output_path}")
