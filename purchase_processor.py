import pandas as pd
import numpy as np
import re
import json
from google import genai
from google.genai import types

import os
API_KEY = os.environ.get("GEMINI_API_KEY")

CA_FIRMS = [
    "BHARATBIZZ VENTURES PRIVATE LIMITED", "STARTUP CLUB INDIA", 
    "LEGIANCE CONSULTANTS PRIVATE LIMITED", "S KRISHAN & ASSOCIATES"
]

def get_rounded_rate(tax_amt, taxable_val):
    if taxable_val == 0 or tax_amt == 0:
        return ''
    raw_rate = (tax_amt / taxable_val) * 100
    valid_rates = [2.5, 5, 6, 9, 12, 14, 18, 20, 28, 40]
    closest_rate = min(valid_rates, key=lambda x: abs(x - raw_rate))
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
        if isinstance(date_val, str) and '-' in date_val:
            parts = date_val.split('-')
            if len(parts[0]) == 4: # yyyy-mm-dd
                return pd.to_datetime(date_val).strftime('%d-%m-%Y')
        return pd.to_datetime(date_val, dayfirst=True).strftime('%d-%m-%Y')
    except:
        return str(date_val)

def find_header_row(file_path, sheet_name, expected_col):
    df_raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None, nrows=10)
    for i, row in df_raw.iterrows():
        if any(expected_col.lower() in str(val).lower() for val in row.values if pd.notna(val)):
            return i
    return 0

def classify_suppliers_with_ai(suppliers, moa_text=None):
    """Uses Gemini API to classify suppliers into appropriate ledgers and groups."""
    if not suppliers:
        return {}
    
    print(f"Calling Gemini AI to classify {len(suppliers)} suppliers...")
    
    moa_context = ""
    if moa_text:
        moa_context = f"\nBusiness Model / MOA Context:\n{moa_text}\n\nUse this context to accurately differentiate between direct 'Purchases' (core to this specific business) and 'Indirect Expenses' (administrative/general expenses).\n"

    prompt = f"""
    You are a Company Law expert and CA.
    Analyze the following supplier names from a business's GSTR2B. 
    Determine the likely business model/expense type for each supplier.{moa_context}
    
    Rules for output:
    1. "Purchase Ledger Name": For core business purchases, use exactly "Purchases". For indirect expenses, determine the specific expense account (e.g., "Legal Fees", "Office Supplies", "Internet Expenses").
    2. "Purchase Group": Must be exactly "Purchase Accounts" (for core purchases) or "Indirect Expenses" (for everything else).
    3. "Party Group": Must be exactly "Sundry Creditors" (if normal business purchases) or "Current Liabilities" (if professional/others, as per Company Law).
    
    Provide the response strictly as a JSON object where keys are the Supplier Names and values are objects containing the 3 fields.
    Example:
    {{
      "BHARTI AIRTEL LTD": {{
         "Purchase Ledger Name": "Telephone & Internet Expenses",
         "Purchase Group": "Indirect Expenses",
         "Party Group": "Current Liabilities"
      }},
      "SOME CORE SUPPLIER": {{
         "Purchase Ledger Name": "Purchases",
         "Purchase Group": "Purchase Accounts",
         "Party Group": "Sundry Creditors"
      }}
    }}
    """
    
    prompt += "\nSuppliers to classify:\n" + json.dumps(suppliers)
    
    try:
        client = genai.Client(api_key=os.environ.get("GEMINI_API_KEY"))
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=prompt,
            config=types.GenerateContentConfig(
                response_mime_type="application/json",
            ),
        )
        result = json.loads(response.text)
        return result
    except Exception as e:
        print(f"Error calling Gemini AI: {e}")
        return {}

def process_purchase_data(gstr2b_path, moa_text=None):
    final_rows = []
    supplier_names = set()
    
    # 1. First pass to extract all supplier names for AI
    try:
        xl = pd.ExcelFile(gstr2b_path)
        sheet_names = xl.sheet_names
        b2b_sheet = next((s for s in sheet_names if any(x in s.lower() for x in ['b2b', 'invoice', 'invoices'])), None)
        
        if b2b_sheet:
            df_b2b_raw = pd.read_excel(gstr2b_path, sheet_name=b2b_sheet, header=None)
            party_col_idx = get_col_index(df_b2b_raw, ['Trade/Legal name', 'Trade/Legal Name', 'Recipients Name', 'Supplier Name'])
            h_idx_b2b = find_header_row(gstr2b_path, b2b_sheet, 'Invoice number')
            if h_idx_b2b == 0:
                h_idx_b2b = find_header_row(gstr2b_path, b2b_sheet, 'Recipients Name')
                
            if party_col_idx != -1:
                supplier_names.update(df_b2b_raw.iloc[h_idx_b2b+1:, party_col_idx].dropna().unique().tolist())
    except Exception as e:
        pass
        
    try:
        cdnr_sheet = next((s for s in sheet_names if any(x in s.lower() for x in ['cdnr', 'note', 'credit note'])), None)
        if cdnr_sheet:
            df_cdnr_raw = pd.read_excel(gstr2b_path, sheet_name=cdnr_sheet, header=None)
            party_col_idx = get_col_index(df_cdnr_raw, ['Trade/Legal name', 'Trade/Legal Name', 'Recipients Name', 'Supplier Name'])
            h_idx_cdnr = find_header_row(gstr2b_path, cdnr_sheet, 'Note number')
            if h_idx_cdnr == 0:
                h_idx_cdnr = find_header_row(gstr2b_path, cdnr_sheet, 'Recipients Name')
                
            if party_col_idx != -1:
                supplier_names.update(df_cdnr_raw.iloc[h_idx_cdnr+1:, party_col_idx].dropna().unique().tolist())
    except:
        pass
        
    supplier_list = [str(s).strip() for s in list(supplier_names) if str(s).strip() and str(s).lower() != 'nan']
    
    # 2. Get AI Classification
    ai_mappings = classify_suppliers_with_ai(supplier_list, moa_text)
    print("AI Mapping complete.")
    
    # 3. Process B2B
    try:
        if b2b_sheet and 'df_b2b_raw' in locals():
            party_col = get_col_index(df_b2b_raw, ['Trade/Legal name', 'Trade/Legal Name', 'Recipients Name', 'Supplier Name', 'Party Name'])
            taxable_col = get_col_index(df_b2b_raw, ['Taxable Value (\u20b9)', 'Taxable Value ()', 'Taxable Value', 'Taxable value'])
            cgst_col = get_col_index(df_b2b_raw, ['Central Tax(\u20b9)', 'Central Tax()', 'Central Tax', 'CGST'])
            sgst_col = get_col_index(df_b2b_raw, ['State/UT Tax(\u20b9)', 'State/UT Tax()', 'State/UT Tax', 'SGST'])
            igst_col = get_col_index(df_b2b_raw, ['Integrated Tax(\u20b9)', 'Integrated Tax()', 'Integrated Tax', 'IGST'])
            inv_num_col = get_col_index(df_b2b_raw, ['Invoice number', 'Invoice No', 'Invoice'])
            inv_date_col = get_col_index(df_b2b_raw, ['Invoice Date', 'Invoice date', 'Date'])
            
            # Safely get header index if not defined
            if 'h_idx_b2b' not in locals():
                h_idx_b2b = find_header_row(gstr2b_path, b2b_sheet, 'Invoice number')
                if h_idx_b2b == 0:
                    h_idx_b2b = find_header_row(gstr2b_path, b2b_sheet, 'Supplier Name')
            
            df_b2b = df_b2b_raw.iloc[h_idx_b2b+1:]
            if inv_num_col != -1:
                df_b2b = df_b2b.dropna(subset=[inv_num_col])
                
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
                
                clean_p = str(party_name).upper().strip()
                if any(ca_firm in clean_p for ca_firm in CA_FIRMS):
                    purchase_ledger = "Legal & Professional Fees"
                    purchase_group = "Indirect Expenses"
                    party_group = "Current Liabilities"
                else:
                    mapping = ai_mappings.get(party_name, {})
                    purchase_ledger = mapping.get('Purchase Ledger Name', 'Purchases')
                    purchase_group = mapping.get('Purchase Group', 'Purchase Accounts')
                    party_group = mapping.get('Party Group', 'Sundry Creditors')

                if party_group == "Current Liabilities" and purchase_group == "Indirect Expenses":
                    voucher_type = "Journal"
                elif party_group == "Sundry Creditors":
                    voucher_type = "Purchase"
                else:
                    voucher_type = "Journal"

                final_rows.append({
                    'Date': format_date(inv_date),
                    'Voucher Type': voucher_type,
                    'Voucher Number': inv_num,
                    'Party Ledger': party_name,
                    'Purchase Ledger': purchase_ledger,
                    'Purchase Item Name': '',
                    'Purchase Item Qty': '',
                    'Purchase Item UOM': '',
                    'Purchase Item Rate': '',
                    'Taxable Value': taxable_val,
                    'CGST Ledger': 'Input CGST' if cgst_amt > 0 else '',
                    'CGST Rate': get_rounded_rate(cgst_amt, taxable_val),
                    'SGST Ledger': 'Input SGST' if sgst_amt > 0 else '',
                    'SGST Rate': get_rounded_rate(sgst_amt, taxable_val),
                    'IGST Ledger': 'Input IGST' if igst_amt > 0 else '',
                    'IGST Rate': get_rounded_rate(igst_amt, taxable_val),
                    'Narration': f"Being purchase made from {party_name}",
                })
    except Exception as e:
        raise Exception(f"Error processing B2B Purchases: {e}")

    # 4. Process B2B-CDNR
    try:
        if 'cdnr_sheet' in locals() and cdnr_sheet and 'df_cdnr_raw' in locals():
            party_col = get_col_index(df_cdnr_raw, ['Trade/Legal name', 'Trade/Legal Name', 'Recipients Name', 'Supplier Name', 'Party Name'])
            taxable_col = get_col_index(df_cdnr_raw, ['Taxable Value (\u20b9)', 'Taxable Value ()', 'Taxable Value', 'Taxable value', 'Note Value'])
            cgst_col = get_col_index(df_cdnr_raw, ['Central Tax(\u20b9)', 'Central Tax()', 'Central Tax', 'CGST'])
            sgst_col = get_col_index(df_cdnr_raw, ['State/UT Tax(\u20b9)', 'State/UT Tax()', 'State/UT Tax', 'SGST'])
            igst_col = get_col_index(df_cdnr_raw, ['Integrated Tax(\u20b9)', 'Integrated Tax()', 'Integrated Tax', 'IGST'])
            note_num_col = get_col_index(df_cdnr_raw, ['Note number', 'Note No', 'Note'])
            note_date_col = get_col_index(df_cdnr_raw, ['Note date', 'Note Date', 'Date'])
            note_type_col = get_col_index(df_cdnr_raw, ['Note type', 'Note Type'])
            
            if 'h_idx_cdnr' not in locals():
                h_idx_cdnr = find_header_row(gstr2b_path, cdnr_sheet, 'Note number')
                if h_idx_cdnr == 0:
                    h_idx_cdnr = find_header_row(gstr2b_path, cdnr_sheet, 'Supplier Name')

            df_cdnr = df_cdnr_raw.iloc[h_idx_cdnr+1:]
            if note_num_col != -1:
                df_cdnr = df_cdnr.dropna(subset=[note_num_col])
                
            for _, row in df_cdnr.iterrows():
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
                note_type = str(row[note_type_col]).strip().upper() if note_type_col != -1 and pd.notna(row[note_type_col]) else ''
                if note_type == 'C':
                    voucher_type = 'Debit Note'
                    narration = f"Being purchase return / debit note for {party_name}"
                elif note_type == 'D':
                    voucher_type = 'Credit Note'
                    narration = f"Being purchase value increase / credit note from {party_name}"
                else:
                    voucher_type = 'Journal'
                    narration = f"Being adjustment for {party_name}"

                clean_p = str(party_name).upper().strip()
                if any(ca_firm in clean_p for ca_firm in CA_FIRMS):
                    purchase_ledger = "Legal & Professional Fees"
                else:
                    mapping = ai_mappings.get(party_name, {})
                    purchase_ledger = mapping.get('Purchase Ledger Name', 'Purchases')

                final_rows.append({
                    'Date': format_date(note_date),
                    'Voucher Type': voucher_type,
                    'Voucher Number': note_num,
                    'Party Ledger': party_name,
                    'Purchase Ledger': purchase_ledger,
                    'Purchase Item Name': '',
                    'Purchase Item Qty': '',
                    'Purchase Item UOM': '',
                    'Purchase Item Rate': '',
                    'Taxable Value': taxable_val,
                    'CGST Ledger': 'Input CGST' if cgst_amt > 0 else '',
                    'CGST Rate': get_rounded_rate(cgst_amt, taxable_val),
                    'SGST Ledger': 'Input SGST' if sgst_amt > 0 else '',
                    'SGST Rate': get_rounded_rate(sgst_amt, taxable_val),
                    'IGST Ledger': 'Input IGST' if igst_amt > 0 else '',
                    'IGST Rate': get_rounded_rate(igst_amt, taxable_val),
                    'Narration': narration,
                })
    except Exception as e:
        raise Exception(f"Error processing Purchase CDNR: {e}")

    # 5. Build Result DataFrame
    cols = ['Date', 'Voucher Type', 'Voucher Number', 'Party Ledger', 'Purchase Ledger', 
            'Purchase Item Name', 'Purchase Item Qty', 'Purchase Item UOM', 'Purchase Item Rate', 
            'Taxable Value', 'CGST Ledger', 'CGST Rate', 'SGST Ledger', 'SGST Rate', 
            'IGST Ledger', 'IGST Rate', 'Narration']
            
    df_final = pd.DataFrame(final_rows, columns=cols)
    
    # 6. Build LEDGER_GROUP_MAP
    ledger_map = []
    
    unique_parties = df_final['Party Ledger'].dropna().unique()
    for party in unique_parties:
        if not party: continue
        mapping = ai_mappings.get(party, {})
        group = mapping.get('Party Group', 'Sundry Creditors')
        ledger_map.append({'Ledger Name': party, 'Under Group': group})
        
    unique_purchases = df_final['Purchase Ledger'].dropna().unique()
    for p_ledger in unique_purchases:
        if not p_ledger: continue
        group = 'Indirect Expenses' # Default
        for k, v in ai_mappings.items():
            if v.get('Purchase Ledger Name') == p_ledger:
                group = v.get('Purchase Group', 'Indirect Expenses')
                break
        ledger_map.append({'Ledger Name': p_ledger, 'Under Group': group})
        
    for tax_ledger in ['Input CGST', 'Input SGST', 'Input IGST']:
        ledger_map.append({'Ledger Name': tax_ledger, 'Under Group': 'Duties & Taxes'})
        
    df_ledgers = pd.DataFrame(ledger_map).drop_duplicates(subset=['Ledger Name'])
    
    return df_final, df_ledgers

if __name__ == "__main__":
    gstr2b_file = r"C:\Users\Admin\.gemini\antigravity\scratch\tally_automation\Input\April May GSTR2B.xlsx"
    
    infinitiminds_moa = "To establish the business of consultancy in the field of software development and to provide services of project management, IT requirements management, recruitment, placement services for IT and others related professionals as per needs of Clients whether domestic or overseas. Main Object of Information Technology Consultancy Company."
    
    result_df, ledgers_df = process_purchase_data(gstr2b_file, moa_text=infinitiminds_moa)
    print("--- Final Processed Purchases (First 5 rows) ---")
    print(result_df.head().to_string())
    
    output_path = r"C:\Users\Admin\.gemini\antigravity\scratch\tally_automation\Output\processed_purchases.xlsx"
    with pd.ExcelWriter(output_path) as writer:
        result_df.to_excel(writer, sheet_name='RAW_DATA_MASTER', index=False)
        ledgers_df.to_excel(writer, sheet_name='LEDGER_GROUP_MAP', index=False)
        
    print(f"Saved updated test output to {output_path}")
