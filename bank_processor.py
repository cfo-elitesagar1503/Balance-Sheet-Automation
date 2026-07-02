import pandas as pd
import numpy as np
import re
import json
import os
from google import genai
from google.genai import types
import openpyxl
from openpyxl.styles import PatternFill

import os
API_KEY = os.environ.get("GEMINI_API_KEY")

CA_FIRMS = {
    "BHARATBIZ": "Bharatbizz Ventures Private Limited",
    "STARTUP CLUB": "Startup Club India", 
    "LEGIANCE": "Legiance Consultants",
    "S KRISHAN": "S Krishan & Associates"
}

GATEWAYS = [
    "PAYU", "RAZORPAY", "CCAVENUE", "BILLDESK", "AMAZON", "FLIPKART", "MEESHO", "STRIPE", "PAYTM"
]

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

COMMON_LEDGERS = [
    "Salary", "Office Expense", "Commission", "Staff Welfare", "Fuel Expense", 
    "Bank Charges", "Internet Expense", "Mobile Recharge Expense", 
    "Legal & Professional Fees", "Rent", "Software Subscription Expense", 
    "Travelling Expense", "Conveyance Expense"
]

BANK_LEDGER_NAME = "Bank Account" # Placeholder for the actual bank ledger name

def clean_narration(text):
    """Deterministically cleans bank narrations."""
    if not isinstance(text, str):
        return ""
    text = text.upper()
    
    # Strip common prefixes
    text = re.sub(r'^(UPI|NEFT|IMPS|RTGS|MMT/IMPS|INF/NEFT)[/-]+', '', text)
    
    # Try to extract party names between slashes for UPI/IMPS
    parts = text.split('/')
    clean_parts = []
    for p in parts:
        p = p.strip()
        # Ignore numeric IDs and bank codes
        if not re.match(r'^[A-Z0-9]+$', p) and len(p) > 2:
            clean_parts.append(p)
            
    if clean_parts:
        return " ".join(clean_parts)
    return text

STOP_WORDS = ["PRIVATE", "LIMITED", "PVT", "LTD", "(I)", "ENTERPRISES", "TRADING", "SOLUTIONS", "TECHNOLOGIES", "TECHCOM", "VENTURES", "SERVICES", "INDIA", "GLOBAL", "CORP", "CORPORATION", "INC", "LLP", "COMPANY", "CO", "AND", "ASSOCIATES", "INDUSTRIES", "FASHION", "APPARELS", "GARMENTS"]

def get_party_keyword(party_name):
    # Remove special chars
    clean_name = re.sub(r'[^A-Z0-9 ]', ' ', party_name.upper())
    words = clean_name.split()
    clean_words = [w for w in words if w not in STOP_WORDS and len(w) > 3]
    if clean_words:
        return clean_words[0]
    return ""

def build_keyword_map(invoices):
    kw_map = {}
    for inv in invoices:
        kw = get_party_keyword(inv['party'])
        if kw:
            if kw not in kw_map:
                kw_map[kw] = set()
            kw_map[kw].add(inv['party'])
    # Return mapping only if keyword uniquely identifies ONE party
    return {kw: list(parties)[0] for kw, parties in kw_map.items() if len(parties) == 1}

def load_gst_masters(sales_path, purchase_path):
    """Loads known debtors and creditors and their invoice amounts from GST outputs."""
    debtors_invoices = []
    creditors_invoices = []
    
    if os.path.exists(sales_path):
        try:
            df_sales = pd.read_excel(sales_path, sheet_name='RAW_DATA_MASTER')
            for _, row in df_sales.iterrows():
                party = str(row.get('Party Ledger', '')).strip()
                date_val = row.get('Date', '')
                taxable = float(row.get('Taxable Value', 0)) if pd.notna(row.get('Taxable Value')) else 0
                cgst = float(row.get('CGST Rate', 0)) if pd.notna(row.get('CGST Rate')) and row.get('CGST Rate') != '' else 0
                sgst = float(row.get('SGST Rate', 0)) if pd.notna(row.get('SGST Rate')) and row.get('SGST Rate') != '' else 0
                igst = float(row.get('IGST Rate', 0)) if pd.notna(row.get('IGST Rate')) and row.get('IGST Rate') != '' else 0
                
                gst_amount = taxable * ((cgst + sgst + igst) / 100.0)
                invoice_amt = taxable + gst_amount
                
                if party and invoice_amt > 0:
                    debtors_invoices.append({
                        'party': party, 
                        'date': date_val, 
                        'taxable': taxable,
                        'gst': gst_amount,
                        'total': invoice_amt,
                        'invoice_no': row.get('Voucher Number', ''),
                        'matched': False
                    })
        except Exception as e:
            print(f"Error loading sales: {e}")
            
    if os.path.exists(purchase_path):
        try:
            df_purch = pd.read_excel(purchase_path, sheet_name='RAW_DATA_MASTER')
            for _, row in df_purch.iterrows():
                party = str(row.get('Party Ledger', '')).strip()
                date_val = row.get('Date', '')
                taxable = float(row.get('Taxable Value', 0)) if pd.notna(row.get('Taxable Value')) else 0
                cgst = float(row.get('CGST Rate', 0)) if pd.notna(row.get('CGST Rate')) and row.get('CGST Rate') != '' else 0
                sgst = float(row.get('SGST Rate', 0)) if pd.notna(row.get('SGST Rate')) and row.get('SGST Rate') != '' else 0
                igst = float(row.get('IGST Rate', 0)) if pd.notna(row.get('IGST Rate')) and row.get('IGST Rate') != '' else 0
                
                gst_amount = taxable * ((cgst + sgst + igst) / 100.0)
                invoice_amt = taxable + gst_amount
                
                if party and invoice_amt > 0:
                    creditors_invoices.append({
                        'party': party, 
                        'date': date_val, 
                        'taxable': taxable,
                        'gst': gst_amount,
                        'total': invoice_amt,
                        'invoice_no': row.get('Voucher Number', ''),
                        'matched': False
                    })
        except Exception as e:
            print(f"Error loading purchases: {e}")
            
    return debtors_invoices, creditors_invoices

def classify_with_ai(transactions, debtors_invoices, creditors_invoices):
    if not transactions:
        return {}
        
    print(f"Calling Gemini AI to classify {len(transactions)} bank transactions...")
    
    unique_debtors = list(set([inv['party'] for inv in debtors_invoices]))
    unique_creditors = list(set([inv['party'] for inv in creditors_invoices]))
    
    ca_firms_list = list(CA_FIRMS.values())
    
    prompt = f"""
    You are an expert Accountant.
    Analyze the following bank transactions (Description, Amount, Cr/Dr).
    
    Your goal is to determine the "Party Name" and "Nature of Expense/Income".
    
    Rules:
    1. For "Party Name": Extract the clearest business or person name from the description.
    2. Check if the Party Name closely matches any of these known Debtors: {unique_debtors}
    3. Check if the Party Name closely matches any of these known Creditors: {unique_creditors}
    4. Check if the Party Name closely matches any of these known CA Firms: {ca_firms_list}. If it does, set the Party Name to the exact CA Firm name and set "Nature of Expense/Income" to "Legal & Professional Fees".
    5. If the transaction matches a known Debtor or Creditor, leave "Nature of Expense/Income" EMPTY (because it's already accounted for in GST).
    6. If the transaction does NOT match a known GST party or CA firm, assign "Nature of Expense/Income" using one of the following preferred ledgers:
       {COMMON_LEDGERS}
       (If none fit, you may use a standard accounting ledger).
       
    Provide the response as a JSON object where keys are the Original Descriptions, and values are objects with "Party Name" and "Nature of Expense/Income".
    """
    
    prompt += "\nTransactions:\n" + json.dumps(transactions)
    
    try:
        client = genai.Client(api_key=os.environ.get("GEMINI_API_KEY"))
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=prompt,
            config=types.GenerateContentConfig(
                response_mime_type="application/json",
            ),
        )
        return json.loads(response.text)
    except Exception as e:
        print(f"Error calling Gemini AI: {e}")
        return {}

def find_matching_invoice(amt, bank_date, invoices, allow_tds=False):
    """Finds a matching invoice checking full amount, taxable amount, gst amount, and optionally TDS variations."""
    target_amt = round(amt)
    tds_rates = [1, 2, 5, 10]
    
    matches = []
    
    for inv in invoices:
        if inv.get('matched'):
            continue
        total = inv['total']
        taxable = inv['taxable']
        gst = inv['gst']
        
        # Scenario 1: Exact Total Match
        if abs(total - target_amt) <= 1.0:
            matches.append((inv, None, "Full Invoice Amount"))
            continue
            
        # Scenario 2: Taxable Only Match
        if abs(taxable - target_amt) <= 1.0:
            matches.append((inv, None, "Taxable Amount Only"))
            continue
            
        # Scenario 3: GST Only Match
        if abs(gst - target_amt) <= 1.0:
            matches.append((inv, None, "GST Amount Only"))
            continue
            
        # Check TDS scenarios (TDS is calculated on Taxable Value)
        if allow_tds:
            matched_tds = False
            for rate in tds_rates:
                tds_deduction = taxable * (rate / 100.0)
                
                # Scenario 4: Full Invoice - TDS
                if abs((total - tds_deduction) - target_amt) <= 1.0:
                    matches.append((inv, rate, f"Full Invoice - {rate}% TDS"))
                    matched_tds = True
                    break
                    
                # Scenario 5: Taxable Only - TDS
                if abs((taxable - tds_deduction) - target_amt) <= 1.0:
                    matches.append((inv, rate, f"Taxable Only - {rate}% TDS"))
                    matched_tds = True
                    break
                    
            if matched_tds:
                continue
            
    if not matches:
        return None, None, None, None
        
    try:
        b_date = pd.to_datetime(bank_date, dayfirst=True)
    except:
        b_date = None
        
    def get_date_diff(match_tuple):
        if b_date is not None:
            try:
                i_date = pd.to_datetime(match_tuple[0]['date'], dayfirst=True)
                return abs((b_date - i_date).days)
            except:
                pass
        return 9999

    matches.sort(key=get_date_diff)
    best_match = matches[0]
    return best_match[0], best_match[1], best_match[2], get_date_diff(best_match)

def find_gateway_matching_invoice(amt, bank_date, invoices, fixed_pct):
    """Finds a matching invoice assuming a strict gateway deduction."""
    target_amt = float(amt)
    if target_amt <= 0:
        return None, None
        
    expected_invoice_amt = target_amt / (1 - fixed_pct / 100.0)
    
    matches = []
    for inv in invoices:
        if inv.get('matched'):
            continue
            
        total = inv['total']
        if total <= 0:
            continue
            
        if abs(total - expected_invoice_amt) <= 1.5:
            matches.append(inv)
            
    if not matches:
        return None, None
        
    # Tie breaker by date
    try:
        if pd.notna(bank_date):
            b_date = pd.to_datetime(bank_date, dayfirst=True)
            def date_diff(inv):
                try:
                    i_date = pd.to_datetime(inv['date'], dayfirst=True)
                    return abs((b_date - i_date).days)
                except:
                    return 9999
            matches.sort(key=date_diff)
    except:
        pass
        
    matches[0]['matched'] = True
    return matches[0]['party'], fixed_pct


def process_bank_statement(bank_path, sales_path, purchase_path, output_path=r"C:\Users\Admin\.gemini\antigravity\scratch\tally_automation\Output\processed_bank_v2.xlsx"):
    df_bank = pd.read_excel(bank_path)
    df_bank.columns = [str(c).strip() for c in df_bank.columns]
    
    date_col = 'Value Date' if 'Value Date' in df_bank.columns else ('Date' if 'Date' in df_bank.columns else df_bank.columns[0])
    desc_col = 'Description' if 'Description' in df_bank.columns else ('Particulars' if 'Particulars' in df_bank.columns else None)
    
    # Add standardized columns
    df_bank['Parsed_Date'] = df_bank[date_col]
    df_bank['Parsed_Desc'] = df_bank[desc_col] if desc_col else ''
    df_bank['Parsed_Amt'] = 0.0
    df_bank['Parsed_Type'] = ''
    
    for idx, row in df_bank.iterrows():
        if 'Withdrawals' in df_bank.columns and 'Deposits' in df_bank.columns:
            w = row.get('Withdrawals', 0)
            d = row.get('Deposits', 0)
            if pd.notna(d) and str(d).strip() != '' and float(d) > 0:
                df_bank.at[idx, 'Parsed_Amt'] = float(d)
                df_bank.at[idx, 'Parsed_Type'] = 'CR'
            elif pd.notna(w) and str(w).strip() != '' and float(w) > 0:
                df_bank.at[idx, 'Parsed_Amt'] = float(w)
                df_bank.at[idx, 'Parsed_Type'] = 'DR'
        else:
            df_bank.at[idx, 'Parsed_Amt'] = float(row.get('Transaction Amount(INR)', 0)) if pd.notna(row.get('Transaction Amount(INR)')) else 0.0
            df_bank.at[idx, 'Parsed_Type'] = str(row.get('Cr/Dr', ''))
            
    # Clean up Parsed_Type
    df_bank['Parsed_Type'] = df_bank['Parsed_Type'].str.upper().str.strip()
    df_bank.loc[df_bank['Parsed_Type'] == 'CREDIT', 'Parsed_Type'] = 'CR'
    df_bank.loc[df_bank['Parsed_Type'] == 'DEBIT', 'Parsed_Type'] = 'DR'
    
    debtors_invoices, creditors_invoices = load_gst_masters(sales_path, purchase_path)
    
    debtor_keywords = build_keyword_map(debtors_invoices)
    creditor_keywords = build_keyword_map(creditors_invoices)
    
    # We will store processed info for each row to build the final columns
    row_info = [{'party': '', 'nature': '', 'remark': '', 'confidence': '', 'clean_desc': '', 'amt': 0.0, 'type': '', 'date': '', 'is_gateway': False} for _ in range(len(df_bank))]
    
    # Pass 1: Pre-process Deterministic
    unmatched_indices = []
    
    for idx, row in df_bank.iterrows():
        desc = str(row.get('Parsed_Desc', ''))
        amt = float(row.get('Parsed_Amt', 0))
        c_d = str(row.get('Parsed_Type', ''))
        b_date = row.get('Parsed_Date', '')
        
        row_info[idx]['amt'] = amt
        row_info[idx]['type'] = c_d
        row_info[idx]['date'] = b_date
        
        if pd.isna(desc) or desc.strip() == '' or desc.lower() == 'description':
            continue
            
        desc_upper = desc.upper()
        clean_desc = clean_narration(desc)
        row_info[idx]['clean_desc'] = clean_desc
        
        # Check if it's a gateway
        is_gateway = any(gw in desc_upper for gw in GATEWAYS)
        row_info[idx]['is_gateway'] = is_gateway
        
        # 1. Deterministic Cash / Contra Logic
        is_cash_entry = 'CASH WDL' in desc_upper or 'CASH DEP' in desc_upper or 'ATM WDL' in desc_upper or ('RVSL' in desc_upper and 'CASH' in desc_upper)
        if is_cash_entry:
            row_info[idx]['party'] = BANK_LEDGER_NAME
            row_info[idx]['nature'] = "Cash"
            row_info[idx]['remark'] = "Matched deterministically as Cash/Contra"
            row_info[idx]['confidence'] = "High"
            continue
            
        # 2. CA Firm Logic
        is_ca_firm = False
        if c_d == 'DR':
            for ca_key, ca_name in CA_FIRMS.items():
                if ca_key in desc_upper:
                    row_info[idx]['party'] = ca_name
                    row_info[idx]['nature'] = "Legal & Professional Fees"
                    row_info[idx]['remark'] = "Matched deterministically as CA Firm"
                    row_info[idx]['confidence'] = "High-Cyan" # Text Only
                    is_ca_firm = True
                    break
        if is_ca_firm:
            continue
            
        unmatched_indices.append(idx)
        
    def check_name_match(party_name, text):
        party_clean = party_name.upper().replace('PRIVATE LIMITED', '').replace('PVT LTD', '').replace(' LTD', '').replace('(I)', '').strip()
        for word in party_clean.split():
            if len(word) > 3 and word in text:
                return True
        return False

    # Pass 2: Strict Narration Name Match
    still_unmatched_indices = []
    for idx in unmatched_indices:
        info = row_info[idx]
        desc_upper = df_bank.loc[idx, 'Parsed_Desc'].upper()
        c_d = info['type']
        amt = info['amt']
        b_date = info['date']
        
        matched = False
        target_dict = debtor_keywords if c_d == 'CR' else creditor_keywords
        target_invoices = debtors_invoices if c_d == 'CR' else creditors_invoices
        
        for kw, party_name in target_dict.items():
            if kw in desc_upper:
                # We found a strict text match! 
                matched = True
                row_info[idx]['party'] = party_name
                
                # Check if amount matches perfectly to catch TDS
                party_invs = [inv for inv in target_invoices if inv['party'] == party_name]
                matched_inv, tds_rate, match_type, date_diff = find_matching_invoice(amt, b_date, party_invs, allow_tds=True)
                
                if matched_inv:
                    matched_inv['matched'] = True
                    row_info[idx]['nature'] = f"TDS Deducted @ {tds_rate}%" if tds_rate else ""
                    row_info[idx]['remark'] = f"Name & Amount Matched ({match_type})"
                    row_info[idx]['confidence'] = "High-Green" # Text + Math
                else:
                    # Partial / On-account payment
                    row_info[idx]['remark'] = "Name Matched via Narration (Partial Amount)"
                    row_info[idx]['confidence'] = "High-Cyan" # Text Only
                break
                
        if not matched:
            still_unmatched_indices.append(idx)
            
    unmatched_indices = still_unmatched_indices

    # Pass 1.5: Strict Invoice Number + Amount Match
    still_unmatched_indices = []
    
    # Matches words like INV, INVOICE, BILL, REF followed by optional chars then number, OR standalone numbers of at least 2 digits
    inv_pattern = re.compile(r'(?:INV(?:OICE)?|BILL|REF)[^\w]*([a-zA-Z0-9\-/]+)|(?<!\d)(\d{2,})(?!\d)', re.IGNORECASE)
    
    for idx in unmatched_indices:
        info = row_info[idx]
        desc = df_bank.loc[idx, 'Parsed_Desc']
        c_d = info['type']
        amt = info['amt']
        b_date = info['date']
        
        target_invoices = debtors_invoices if c_d == 'CR' else creditors_invoices
        
        matches = inv_pattern.findall(str(desc))
        potential_invs = []
        for m in matches:
            if m[0]: potential_invs.append(str(m[0]).strip().upper())
            if m[1]: potential_invs.append(str(m[1]).strip().upper())
            
        matched = False
        if potential_invs:
            for inv in target_invoices:
                if inv.get('matched'): continue
                inv_no_clean = str(inv.get('invoice_no', '')).strip().upper()
                if not inv_no_clean: continue
                
                is_inv_match = False
                for p_inv in potential_invs:
                    if p_inv in inv_no_clean or inv_no_clean in p_inv:
                        is_inv_match = True
                        break
                        
                if is_inv_match:
                    matched_inv, tds_rate, match_type, date_diff = find_matching_invoice(amt, b_date, [inv], allow_tds=True)
                    if matched_inv:
                        inv['matched'] = True
                        row_info[idx]['party'] = inv['party']
                        row_info[idx]['nature'] = f"TDS Deducted @ {tds_rate}%" if tds_rate else ""
                        row_info[idx]['remark'] = f"Invoice No. & Amount Matched ({match_type})"
                        row_info[idx]['confidence'] = "High-Green"
                        matched = True
                        break
                        
        if not matched:
            still_unmatched_indices.append(idx)
            
    unmatched_indices = still_unmatched_indices

    # Pass 2a: Match Single Exact Amounts (No TDS)
    still_unmatched_indices = []
    for idx in unmatched_indices:
        info = row_info[idx]
        amt = info['amt']
        c_d = info['type']
        b_date = info['date']
        desc_upper = df_bank.loc[idx, 'Parsed_Desc'].upper()
        
        matched = False
        if c_d == 'CR':
            matched_inv, tds_rate, match_type, date_diff = find_matching_invoice(amt, b_date, debtors_invoices, allow_tds=False)
            if matched_inv:
                is_name_match = check_name_match(matched_inv['party'], desc_upper)
                if is_name_match or date_diff <= 45:
                    matched_inv['matched'] = True
                    row_info[idx]['party'] = matched_inv['party']
                    row_info[idx]['nature'] = f"TDS Deducted @ {tds_rate}%" if tds_rate else ""
                    if is_name_match:
                        row_info[idx]['remark'] = f"Amount & Name Matched ({match_type})"
                        row_info[idx]['confidence'] = "High-Green"
                    else:
                        row_info[idx]['remark'] = f"Amount & Date Matched ({match_type})"
                        row_info[idx]['confidence'] = "High-Cyan"
                    matched = True
        elif c_d == 'DR':
            matched_inv, tds_rate, match_type, date_diff = find_matching_invoice(amt, b_date, creditors_invoices, allow_tds=False)
            if matched_inv:
                is_name_match = check_name_match(matched_inv['party'], desc_upper)
                if is_name_match or date_diff <= 45:
                    matched_inv['matched'] = True
                    row_info[idx]['party'] = matched_inv['party']
                    row_info[idx]['nature'] = f"TDS Deducted @ {tds_rate}%" if tds_rate else ""
                    if is_name_match:
                        row_info[idx]['remark'] = f"Amount & Name Matched ({match_type})"
                        row_info[idx]['confidence'] = "High-Green"
                    else:
                        row_info[idx]['remark'] = f"Amount & Date Matched ({match_type})"
                        row_info[idx]['confidence'] = "High-Cyan"
                    matched = True
                
        if not matched:
            still_unmatched_indices.append(idx)

    unmatched_indices = still_unmatched_indices

    # Pass 2b: Match Single Amounts with TDS
    still_unmatched_indices = []
    for idx in unmatched_indices:
        info = row_info[idx]
        amt = info['amt']
        c_d = info['type']
        b_date = info['date']
        desc_upper = df_bank.loc[idx, 'Parsed_Desc'].upper()
        
        matched = False
        if c_d == 'CR':
            matched_inv, tds_rate, match_type, date_diff = find_matching_invoice(amt, b_date, debtors_invoices, allow_tds=True)
            if matched_inv:
                is_name_match = check_name_match(matched_inv['party'], desc_upper)
                if is_name_match or date_diff <= 45:
                    matched_inv['matched'] = True
                    row_info[idx]['party'] = matched_inv['party']
                    row_info[idx]['nature'] = f"TDS Deducted @ {tds_rate}%" if tds_rate else ""
                    if is_name_match:
                        row_info[idx]['remark'] = f"Amount & Name Matched ({match_type})"
                        row_info[idx]['confidence'] = "High-Green"
                    else:
                        row_info[idx]['remark'] = f"Amount & Date Matched ({match_type})"
                        row_info[idx]['confidence'] = "High-Cyan"
                    matched = True
        elif c_d == 'DR':
            matched_inv, tds_rate, match_type, date_diff = find_matching_invoice(amt, b_date, creditors_invoices, allow_tds=True)
            if matched_inv:
                is_name_match = check_name_match(matched_inv['party'], desc_upper)
                if is_name_match or date_diff <= 45:
                    matched_inv['matched'] = True
                    row_info[idx]['party'] = matched_inv['party']
                    row_info[idx]['nature'] = f"TDS Deducted @ {tds_rate}%" if tds_rate else ""
                    if is_name_match:
                        row_info[idx]['remark'] = f"Amount & Name Matched ({match_type})"
                        row_info[idx]['confidence'] = "High-Green"
                    else:
                        row_info[idx]['remark'] = f"Amount & Date Matched ({match_type})"
                        row_info[idx]['confidence'] = "High-Cyan"
                    matched = True
                
        if not matched:
            still_unmatched_indices.append(idx)
            
    unmatched_indices = still_unmatched_indices

    # Pass 3: Grouping Logic for Exact Amounts
    groups = {}
    for idx in unmatched_indices:
        info = row_info[idx]
        key = (info['date'], info['clean_desc'], info['type'])
        if key not in groups:
            groups[key] = []
        groups[key].append(idx)
        
    still_unmatched_indices = []
    for key, indices in groups.items():
        if len(indices) > 1:
            total_amt = sum(row_info[i]['amt'] for i in indices)
            b_date = key[0]
            c_d = key[2]
            
            matched = False
            if c_d == 'CR':
                matched_inv, tds_rate, match_type, date_diff = find_matching_invoice(total_amt, b_date, debtors_invoices, allow_tds=True)
                if matched_inv:
                    matched_inv['matched'] = True
                    for i in indices:
                        row_info[i]['party'] = matched_inv['party']
                        row_info[i]['nature'] = f"TDS Deducted @ {tds_rate}%" if tds_rate else ""
                        row_info[i]['remark'] = f"Grouped Payment Matched ({match_type})"
                        row_info[i]['confidence'] = "High-Cyan"
                    matched = True
            elif c_d == 'DR':
                matched_inv, tds_rate, match_type, date_diff = find_matching_invoice(total_amt, b_date, creditors_invoices, allow_tds=True)
                if matched_inv:
                    matched_inv['matched'] = True
                    for i in indices:
                        row_info[i]['party'] = matched_inv['party']
                        row_info[i]['nature'] = f"TDS Deducted @ {tds_rate}%" if tds_rate else ""
                        row_info[i]['remark'] = f"Grouped Payment Matched ({match_type})"
                        row_info[i]['confidence'] = "High-Cyan"
                    matched = True
                        
            if not matched:
                still_unmatched_indices.extend(indices)
        else:
            still_unmatched_indices.extend(indices)
            
    unmatched_indices = still_unmatched_indices

    # Pass 4: Gateway Matches (Dynamic Uniform Percentage)
    gateway_pcts = {}
    for gw in GATEWAYS:
        pct_counts = {}
        gw_indices = [i for i in unmatched_indices if gw in df_bank.loc[i, 'Parsed_Desc'].upper()]
        for idx in gw_indices:
            info = row_info[idx]
            amt = info['amt']
            invoices = debtors_invoices if info['type'] == 'CR' else creditors_invoices
            for inv in invoices:
                if inv.get('matched'): continue
                if inv['total'] <= 0: continue
                pct = round(((inv['total'] - amt) / inv['total']) * 100, 2)
                if 0.5 <= pct <= 5.0:
                    expected_inv = amt / (1 - pct / 100.0)
                    if abs(inv['total'] - expected_inv) <= 1.5:
                        pct_counts[pct] = pct_counts.get(pct, 0) + 1
        if pct_counts:
            # Find the most common percentage for this gateway
            best_pct = max(pct_counts.items(), key=lambda x: x[1])[0]
            gateway_pcts[gw] = best_pct

    still_unmatched_indices = []
    for idx in unmatched_indices:
        info = row_info[idx]
        matched = False
        if info['is_gateway']:
            desc_upper = df_bank.loc[idx, 'Parsed_Desc'].upper()
            my_gw = next((gw for gw in GATEWAYS if gw in desc_upper), None)
            if my_gw and my_gw in gateway_pcts:
                fixed_pct = gateway_pcts[my_gw]
                invoices = debtors_invoices if info['type'] == 'CR' else creditors_invoices
                matched_party, ded_pct = find_gateway_matching_invoice(info['amt'], info['date'], invoices, fixed_pct)
                if matched_party:
                    row_info[idx]['party'] = matched_party
                    row_info[idx]['remark'] = f"Gateway Payment (exact {ded_pct}% platform deduction)"
                    row_info[idx]['confidence'] = "High-Cyan" # Math Only
                    matched = True
        
        if not matched:
            still_unmatched_indices.append(idx)
            
    # Pass 3: AI Fallback for remaining unmatched
    unique_txns = []
    for idx in still_unmatched_indices:
        info = row_info[idx]
        desc = df_bank.loc[idx, 'Parsed_Desc']
        unique_txns.append({
            'original': desc,
            'clean': info['clean_desc'],
            'amount': info['amt'],
            'type': info['type']
        })
        
    ai_mappings = classify_with_ai(unique_txns, debtors_invoices, creditors_invoices)
    
    for idx in still_unmatched_indices:
        desc = df_bank.loc[idx, 'Parsed_Desc']
        mapping = ai_mappings.get(desc, {})
        party = mapping.get('Party Name', '')
        nature = mapping.get('Nature of Expense/Income', '')
        remark = "Matched by AI (Narration)"
        confidence = "Medium"
        
        # Check if AI suggested a CA firm name (fuzzy match)
        is_ca_firm_ai = False
        if party:
            party_upper = party.upper()
            for ca_key, ca_full_name in CA_FIRMS.items():
                if ca_key in party_upper:
                    party = ca_full_name
                    nature = "Legal & Professional Fees"
                    remark = "Matched by AI (Overridden to CA Firm)"
                    confidence = "High-Cyan" # Treat as high confidence
                    is_ca_firm_ai = True
                    break
                    
        # Upgrade Exotel/Clear narration matches!
        if not is_ca_firm_ai:
            if party:
                desc_upper = desc.upper()
                if check_name_match(party, desc_upper):
                    remark = "Matched by AI (Exact Name in Narration)"
                    confidence = "High-Cyan" # Text Only
            elif not nature:
                remark = "Low Confidence AI Fallback"
                confidence = "Low" # Red
            
        row_info[idx]['party'] = party
        row_info[idx]['nature'] = nature
        row_info[idx]['remark'] = remark
        row_info[idx]['confidence'] = confidence
        
    # Write back to dataframe
    df_bank['Party Name'] = [info['party'] for info in row_info]
    df_bank['Nature of Expense/Income'] = [info['nature'] for info in row_info]
    df_bank['Remarks'] = [info['remark'] for info in row_info]
    df_bank['Confidence'] = [info['confidence'] for info in row_info]
    df_bank['Cr/Dr'] = [info['type'] for info in row_info]
    df_bank[date_col] = df_bank[date_col].apply(format_date)
    
    cols_to_drop = ['Parsed_Date', 'Parsed_Desc', 'Parsed_Amt', 'Parsed_Type']
    df_bank = df_bank.drop(columns=[c for c in cols_to_drop if c in df_bank.columns])
    
    return df_bank

def apply_excel_formatting(output_path):
    wb = openpyxl.load_workbook(output_path)
    ws = wb.active
    
    # Find confidence column index
    headers = [str(cell.value) for cell in ws[1]]
    if 'Confidence' not in headers:
        return
    
    conf_idx = headers.index('Confidence') + 1
    
    green_fill = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    red_fill = PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid")
    
    cyan_fill = PatternFill(start_color="E0FFFF", end_color="E0FFFF", fill_type="solid")
    blue_fill = PatternFill(start_color="C9DAF8", end_color="C9DAF8", fill_type="solid")
    purple_fill = PatternFill(start_color="E4D7F5", end_color="E4D7F5", fill_type="solid")
    orange_fill = PatternFill(start_color="FCE5CD", end_color="FCE5CD", fill_type="solid")
    
    for row in range(2, ws.max_row + 1):
        conf = ws.cell(row=row, column=conf_idx).value
        fill = None
        if conf == 'High' or conf == 'High-Green':
            fill = green_fill
        elif conf == 'High-Cyan':
            fill = cyan_fill
        elif conf == 'High-Blue':
            fill = blue_fill
        elif conf == 'High-Purple':
            fill = purple_fill
        elif conf == 'High-Orange':
            fill = orange_fill
        elif conf == 'Medium':
            fill = yellow_fill
        elif conf == 'Low':
            fill = red_fill
            
        if fill:
            for col in range(1, ws.max_column + 1):
                ws.cell(row=row, column=col).fill = fill
                
    wb.save(output_path)

if __name__ == "__main__":
    bank_file = r"C:\Users\Admin\.gemini\antigravity\scratch\tally_automation\Input\April May Bank.xls"
    sales_file = r"C:\Users\Admin\.gemini\antigravity\scratch\tally_automation\Output\processed_sales.xlsx"
    purchase_file = r"C:\Users\Admin\.gemini\antigravity\scratch\tally_automation\Output\processed_purchases.xlsx"
    
    df_result = process_bank_statement(bank_file, sales_file, purchase_file)
    
    # output_path is provided as a parameter
    df_result.to_excel(output_path, index=False)
    
    # Apply coloring based on Confidence
    apply_excel_formatting(output_path)
    print(f"\nSaved updated bank statement to {output_path} with Confidence highlighting.")
