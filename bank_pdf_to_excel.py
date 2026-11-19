import os
import re
import glob
import sys
import io
from datetime import datetime
import pdfplumber
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# EasyOCR lazy-load singleton  (only initialised when a scanned PDF is found)
# ---------------------------------------------------------------------------
_ocr_reader = None

def get_ocr_reader():
    """Return a shared EasyOCR Reader, loading it on the very first call."""
    global _ocr_reader
    if _ocr_reader is None:
        try:
            import easyocr
            print("[OCR] Initialising EasyOCR (downloading model on first run — please wait)...")
            _ocr_reader = easyocr.Reader(['en'], gpu=False, verbose=False)
            print("[OCR] EasyOCR ready.")
        except ImportError:
            raise ImportError(
                "EasyOCR is not installed. Run:  pip install easyocr pymupdf numpy"
            )
    return _ocr_reader

# Helper function to clean numeric amounts
def clean_amount(val_str):
    if not val_str:
        return 0.0
    val_str = str(val_str).replace(",", "").strip()
    is_negative = False
    
    # Check for negative signs or indicators
    if val_str.startswith("-"):
        is_negative = True
        val_str = val_str[1:]
    elif re.search(r'(?i)\b(?:dr|\(dr\)|dr\.)', val_str):
        is_negative = True
    elif re.search(r'(?i)\b(?:cr|\(cr\)|cr\.)', val_str):
        is_negative = False
        
    # Remove (Cr), (Dr), Cr., Dr., Cr, Dr indicators
    val_str = re.sub(r'\([^\)]*\)', '', val_str)
    val_str = re.sub(r'(?i)\b(?:dr|cr)\b\.?', '', val_str)
    val_str = re.sub(r'(?i)(?:cr|dr)\.?', '', val_str)
    
    # Replace OCR-misread 'O', 'o', 'Q' or 'OO' with '0' (e.g. 4O5O -> 4050, 5OO.OO -> 500.00)
    val_str = val_str.upper().replace("OO", "00").replace("O", "0").replace("Q", "0")
    
    # Remove non-numeric characters except decimal points
    val_str = re.sub(r'[^0-9\.]', '', val_str).strip()
    val_str = val_str.rstrip('.')
    if val_str.count('.') > 1:
        parts = val_str.split('.')
        val_str = parts[0] + '.' + ''.join(parts[1:])
        
    try:
        val = float(val_str)
        return -val if is_negative else val
    except ValueError:
        return 0.0

# Helper function to parse dates to YYYY-MM-DD
def parse_date(date_str):
    if not date_str:
        return ""
    date_str = date_str.strip()
    
    # 1. First try parsing directly if it's already a standard numeric format
    # to avoid unnecessary regex replacements
    for fmt in ('%d/%m/%Y', '%d/%m/%y', '%d-%m-%Y', '%d-%m-%y', '%d.%m.%Y', '%d.%m.%y', '%Y-%m-%d'):
        try:
            dt = datetime.strptime(date_str, fmt)
            return dt.strftime('%Y-%m-%d')
        except ValueError:
            pass
            
    # 2. Normalize alphabetical English months to their 2-digit numeric equivalents
    # ordered from longest name to shortest name to avoid partial replacement issues
    months_map = {
        'september': '09', 'february': '02', 'november': '11', 'december': '12',
        'january': '01', 'october': '10', 'august': '08', 'march': '03',
        'april': '04', 'june': '06', 'july': '07', 'sept': '09',
        'jan': '01', 'feb': '02', 'mar': '03', 'apr': '04', 'may': '05',
        'jun': '06', 'jul': '07', 'aug': '08', 'sep': '09', 'oct': '10',
        'nov': '11', 'dec': '12'
    }
    
    normalized_str = date_str.lower()
    for month_name, month_num in months_map.items():
        # Match only full tokens of the month name (not part of another word)
        normalized_str = re.sub(r'(?<![a-z])' + re.escape(month_name) + r'(?![a-z])', month_num, normalized_str)
        
    # 3. Try parsing the normalized numeric date string
    for fmt in (
        '%d/%m/%Y', '%d/%m/%y', 
        '%d-%m-%Y', '%d-%m-%y',
        '%d.%m.%Y', '%d.%m.%y', 
        '%d %m %Y', '%Y-%m-%d'
    ):
        try:
            dt = datetime.strptime(normalized_str, fmt)
            return dt.strftime('%Y-%m-%d')
        except ValueError:
            pass
            
    # 4. Try whitespace/separator normalization fallback
    try:
        normalized = re.sub(r'[\s\-./]+', ' ', normalized_str)
        dt = datetime.strptime(normalized.strip(), '%d %m %Y')
        return dt.strftime('%Y-%m-%d')
    except ValueError:
        pass
        
    try:
        normalized = re.sub(r'[\s\-./]+', ' ', normalized_str)
        dt = datetime.strptime(normalized.strip(), '%d %m %y')
        return dt.strftime('%Y-%m-%d')
    except ValueError:
        pass
        
    return date_str

# Helper function to filter out header, footer, page number, address, and metadata lines from transaction narration
def is_noise_line(line, bank_type):
    line_clean = line.strip()
    if not line_clean:
        return True
        
    # Transaction lines starting with a date are NEVER noise lines
    if re.match(r'^(?:(\d+)\s+)?\d{2}[/\-\.]\d{2}[/\-\.]\d{2,4}\s+', line_clean):
        return False
        
    line_lower = line_clean.lower()
    
    # Specific page boundary table header artifacts
    if line_lower in ('ref no/', 'cheque', 'no', 'no.', 'ref no/cheque', 'ref no/ cheque', 'value date', 'post date', 'details', '? debit', '? credit', 'balance'):
        return True
        
    # Page numbers
    if re.search(r'\bpage\b', line_lower) or re.search(r'\bpg\b', line_lower):
        return True
        
    # Universal metadata indicators
    noise_indicators = [
        'statement of account', 'account statement', 'registered office', 'customer id',
        'cif id', 'cif no', 'account number', 'account no', 'statement period',
        'opening balance', 'closing balance', 'total debit', 'total credit',
        'txn date', 'value date', 'particulars', 'cheque no', 'debit', 'credit',
        'balance', 'always you first', 'customer name', 'communication address',
        'email id', 'phone no', 'ckyc id', 'nomination', 'nominee name',
        'account branch', 'branch address', 'ifsc', 'micr', 'account opening date',
        'account status', 'account type', 'currency', 'branch code', 'ref no.',
        'cheque id', 'transaction details', 'withdrawal', 'deposit', 'txn. date',
        'post date', 'narration', 'cheque details', 'serial no', 'si no',
        's.no', 'remarks', 'transaction id', 'chq num', 'chq. no.', 'drawing power',
        'interest rate', 'mod balance', 'ifs code', 'micr code', 'balance as on',
        'sol id', 'sol description', 'sol address', 'limit sanctioned', 'sanctioned int. rate',
        'joint/auth. sign.', 'nominated', 'nomination registered', 'branch name',
        'joint holders', 'this is a computer generated statement', 'end of statement', 
        'end of the statement', 'generated on', 'generated date', 're kyc due date', 'customer/cif id',
        'care and cure corpor', 'v n shah and co', 'linked casa accounts', 'linked deposits',
        'linked loan', 'linked lockers', 'other digital products', 'sms alert', 'debit card',
        'internet banking', 'disclaimer:', 'for any support or clarification',
        'please provide your contact', 'contents of this statement', 'unless the constituent',
        'closing total', 'transaction count', 'this is system generated', 'this is a computer generated',
        'important message', 'important safety tips',
        'union bank', 'state bank of india', 'hdfc bank', 'kotak mahindra bank', 'idfc first bank',
        'rbl bank', 'the modasa nagarik', 'bank of india', 'the sabarkantha dccb', 'the mehsana urban',
        'statement summary', 'brought forward', 'carried forward', 'please do not share',
        'please verify your transactions', 'please contact your branch', 'atm, debit/credit card',
        'bank never asks', 'power of attorney', 'does not require a signature',
        'discrepancy', 'note:', 'bank of baroda', 'bob world', 'normal course of business',
        'base branch', 'raise & complaint', 'raise a complaint', 'date number', 'date date number',
        'computer-generated', 'computer generated'
    ]
    if any(ind in line_lower for ind in noise_indicators):
        return True
        
    # Prefix indicators for header metadata
    metadata_starts = (
        'address', 'name', 'pan', 'ckyc', 'product', 'a/c open date', 'a/c status', 'gstin',
        'email', 'phone', 'mobile', 'joint', 'nominee', 'home branch', 'customer name',
        'communication address', 'branch address', 'registered office', 'sol id', 'sol description',
        'cleared balance', 'uncleared amount'
    )
    if line_lower.startswith(metadata_starts):
        return True
        
    # Custom bank-specific headers/addresses details
    if bank_type == "SK":
        if any(x in line_lower for x in ('patidar petroleum', 'govindpur', 'malpur')) and not any(t in line_lower for t in ('dep tfr:', 'wdl tfr:')):
            return True
    elif bank_type == "MEHSANA":
        if any(x in line_lower for x in ('m/s om auto parts', 'shriram city center', 'shamlaji road')):
            return True
    elif bank_type == "UBI1":
        if 'shree ram gruh' in line_lower:
            return True
    elif bank_type == "UBI2":
        if any(x in line_lower for x in ('patel dharmendrakumar', 'sunderpura kampa', 'am lai')):
            return True
    elif bank_type == "UBI3":
        if any(x in line_lower for x in ('amitkumar jayantilal', 'at amaratpura kampa', 'khambhisar,modasa', 'ubin0555207', 'malpur road', 'modasa1.shyam', 'nagarmalpur road', 'opp iti', 'customer/cif id', 'branch address')):
            return True
    elif bank_type == "HDFC":
        if any(x in line_lower for x in (
            'm/s. vaishali enterprise', 'ground floor vadagaam', 'dhansura road', 'dhansura 383307',
            'm/s. sai traders', 'gokulesh complex', 'we understand your world', 'shamalaji 383355',
            'prime potential', 'biz pro plus', 'cust id', 'rtgs/neft ifsc', 'branch code',
            'statement of account', 'registered office address', 'funds earmarked for hold',
            'within 30 days of receipt', 'statement summary', 'opening balance', 'dr count', 'cr count',
            'closing bal', 'generated on:', 'generated by:', 'requesting branch code', 'gstin number details',
            'io hdfc', 'hdfc rank', 'hdfc bank limited', 'page no', 'from :', 'to :'
        )):
            return True
    elif bank_type == "IDFC":
        if any(x in line_lower for x in ('chovatiya tusharbhai', 'c/o chovatiya', 'b-4,shridhar', 'choila road', 'bayad 383325', 'arvalli 383325')):
            return True
    elif bank_type == "ICICI":
        if any(x in line_lower for x in (
            'detailed statement', 'cust id', 'ifsc code: icic', 'statement request/download',
            'amount from: na to na', 'cheque number from: na to na', 'transaction type: dc',
            'legends used in account statement', 'bbps - bharat bill payment', 'bctt - banking cash',
            'bil - internet bill payment', 'bpay - bill payment', 'ccwd - cardless cash',
            'dtax - direct tax', 'eba - transaction on icici', 'idtx - indirect tax',
            'imps - immediate payment', 'inf - internet fund transfer', 'inft - internal fund transfer',
            'lccbrn cms - local cheque', 'lnpy - linked loan payment', 'mmt - mobile money transfer',
            'n chg - neft charges', 'neft - national electronics', 'onl - online shopping',
            'pac - personal accident', 'pavc - pay any visa', 'payc - pay to contact',
            'rchg - recharge', 'smo - smart money order', 't chg - travel charges',
            'top - mobile recharge', 'uccbrn cms - upcountry cheque', 'vat / mat / nfs',
            'vps / ips - debit card', 'gib - tax & statutory', 'cms - internet bulk payment',
            '*you can download maximum of 2000', 'this is a system-generated statement'
        )):
            return True
    elif bank_type == "BOI":
        if any(x in line_lower for x in (
            'bank of india', 'statement of account', 'opening balance:', 'custid :', 'a/c no :',
            'type : savings', 'currency : inr', 'ifsc code : bkid', 'micr code :', 'joint holder:',
            'nominee : yes', 'dhansura', 'bhuj', 'page:'
        )):
            return True

    return False

# Identify bank statement type from first page text
def detect_bank_type(first_page_text):
    text_lower = first_page_text.lower()
    
    # 1. Bank of India (Check before others to prevent UPI /HDFC/ or /BARB/ routing)
    if "bank of india" in text_lower or "bkid" in text_lower or "star token" in text_lower:
        return "BOI"
        
    # 2. ICICI Bank (Check headers/IFSC first to avoid cross-bank false positives from customer names)
    if "icici bank" in text_lower or "icic0" in text_lower or "vlclci" in text_lower or "corporate office: icici" in text_lower or "icicibank.com" in text_lower or "icici direct" in text_lower:
        return "ICICI"
        
    # 3. HDFC Bank (Strict header/IFSC check)
    if "hdfc bank" in text_lower or "hdfc0" in text_lower or "hdfcbank" in text_lower or "hdfc rank" in text_lower or "io hdfc" in text_lower or "hdfc bank limited" in text_lower or "www.hdfcbank.com" in text_lower:
        return "HDFC"
        
    # 4. State Bank of India
    if "sbin" in text_lower or "state bank of india" in text_lower:
        return "SBI"
        
    # 5. Union Bank of India
    if "union bank" in text_lower or "ubin" in text_lower or "patel dharmendrakumar" in text_lower:
        if "patel dharmendrakumar" in text_lower or "ubin0916846" in text_lower:
            return "UBI2"
        return "UBI3"
    elif "shree ram gruh" in text_lower and ("union" in text_lower or "ubin" in text_lower):
        return "UBI1"
        
    # 6. Other banks
    if "kotak" in text_lower or ("kkbk" in text_lower and "mnsb" not in text_lower):
        return "KOTAK"
    elif "idfc" in text_lower or "idfb0" in text_lower:
        return "IDFC"
    elif "rbl bank" in text_lower or "ratn0" in text_lower:
        return "RBL"
    elif "modasa nagarik" in text_lower or "kkbk0mnsb" in text_lower:
        return "MODASA"
    elif "sabarkantha dccb" in text_lower or "gscb0skb" in text_lower:
        return "SK"
    elif "mehsana" in text_lower or "msnu0" in text_lower:
        return "MEHSANA"
    elif "bank of baroda" in text_lower or "barb0" in text_lower or "bob world" in text_lower:
        return "BOB"
    elif "axis bank" in text_lower or "utib0" in text_lower:
        return "AXIS"
    elif "prime bank" in text_lower or "pmec0" in text_lower or "pcbl /" in text_lower or "primebankindia" in text_lower:
        return "PRIME"
        
    return "GENERIC"

def is_sbi_header(line):
    line_clean = line.strip().upper()
    headers = [
        "WDL TFR", "DEP TFR", "DEPTFR", "DEBIT ACH", "CREDIT ACH", 
        "CREDIT CHQ", "DEBIT CHQ", "CHEQUE WDL", "CHQ WDL", 
        "CASH DEPOSIT", "CASH DEP", "INTEREST CREDIT"
    ]
    for h in headers:
        if line_clean.startswith(h):
            return True
    return False

def parse_sbi(pdf_file_obj, lines):
    # 1. Dynamic Table Extraction across all pages (handles 6-col, 7-col, 8-col SBI formats)
    try:
        if hasattr(pdf_file_obj, "seek"):
            pdf_file_obj.seek(0)
        with pdfplumber.open(pdf_file_obj) as pdf:
            txs = []
            col_map = {}
            header_col_count = 0
            
            for page in pdf.pages:
                tables = page.extract_tables()
                if not tables:
                    continue
                for table in tables:
                    if not table or len(table) < 1:
                        continue
                    
                    found_header = False
                    header_row = None
                    
                    for row_idx in range(min(5, len(table))):
                        row = [str(x or "").lower().replace("\n", " ").strip() for x in table[row_idx]]
                        has_date = any('date' in x or 'dt' in x for x in row)
                        has_desc = any('desc' in x or 'particular' in x or 'narration' in x or 'detail' in x for x in row)
                        
                        if has_date and has_desc:
                            header_row = table[row_idx]
                            col_map = {}
                            header_col_count = len(row)
                            for col_idx, col_name in enumerate(row):
                                col_name_clean = col_name.replace(".", " ").replace("(", " ").replace(")", " ").replace("/", " ").replace("-", " ").strip()
                                tokens = col_name_clean.split()
                                if 'date' in col_name_clean or 'dt' in col_name_clean:
                                    if 'val' in col_name_clean or 'value' in col_name_clean:
                                        col_map['val_date'] = col_idx
                                    elif 'txn' in col_name_clean or 'tx' in col_name_clean or 'trans' in col_name_clean or 'date' in col_name_clean:
                                        if 'date' not in col_map:
                                            col_map['date'] = col_idx
                                        else:
                                            col_map['val_date'] = col_idx
                                elif 'desc' in col_name_clean or 'particular' in col_name_clean or 'narration' in col_name_clean or 'detail' in col_name_clean:
                                    col_map['desc'] = col_idx
                                elif 'chq' in col_name_clean or 'cheque' in col_name_clean or 'ref' in col_name_clean:
                                    col_map['ref'] = col_idx
                                elif 'branch' in col_name_clean:
                                    col_map['branch'] = col_idx
                                elif any(x in col_name_clean for x in ('debit', 'withdrawal', 'withdraw', 'dr')) or 'dr' in tokens:
                                    col_map['debit'] = col_idx
                                elif any(x in col_name_clean for x in ('credit', 'deposit', 'cr')) or 'cr' in tokens:
                                    col_map['credit'] = col_idx
                                elif 'balance' in col_name_clean or 'bal' in col_name_clean:
                                    col_map['balance'] = col_idx
                            found_header = True
                            break
                            
                    if not found_header and col_map and len(table[0]) == header_col_count:
                        start_row_idx = 0
                        parse_table = True
                    elif found_header and 'date' in col_map and 'desc' in col_map:
                        header_idx = table.index(header_row)
                        start_row_idx = header_idx + 1
                        parse_table = True
                    else:
                        parse_table = False
                        
                    if parse_table:
                        for r_idx in range(start_row_idx, len(table)):
                            row = table[r_idx]
                            if len(row) < header_col_count:
                                continue
                                
                            raw_date = str(row[col_map['date']] or "").strip()
                            date_m = re.search(r'\b\d{2}[/\-\.]\d{2}[/\-\.]\d{2,4}\b', raw_date)
                            
                            if not date_m:
                                # Continuation line
                                if txs and 'desc' in col_map and row[col_map['desc']]:
                                    desc_cont = str(row[col_map['desc']]).strip().replace("\n", " ")
                                    if desc_cont and not is_noise_line(desc_cont, "SBI") and not any(x in desc_cont.lower() for x in ("total", "carried forward", "brought forward", "page", "statement")):
                                        txs[-1]['Description'] += " " + desc_cont
                                continue
                                
                            date_str = date_m.group(0)
                            val_date_str = ""
                            if 'val_date' in col_map and row[col_map['val_date']]:
                                val_m = re.search(r'\b\d{2}[/\-\.]\d{2}[/\-\.]\d{2,4}\b', str(row[col_map['val_date']]))
                                if val_m:
                                    val_date_str = val_m.group(0)
                            if not val_date_str:
                                val_date_str = date_str
                                
                            desc_val = str(row[col_map['desc']] or "").strip().replace("\n", " ")
                            
                            ref_val = ""
                            if 'ref' in col_map and row[col_map['ref']]:
                                raw_ref = str(row[col_map['ref']]).strip().replace("\n", " ")
                                ref_m = re.search(r'\b\d{5,22}\b', raw_ref)
                                if ref_m:
                                    ref_val = ref_m.group(0)
                                elif raw_ref and raw_ref != "-":
                                    ref_val = raw_ref
                                    
                            debit_val = clean_amount(row[col_map['debit']]) if 'debit' in col_map else 0.0
                            credit_val = clean_amount(row[col_map['credit']]) if 'credit' in col_map else 0.0
                            balance_val = clean_amount(row[col_map['balance']]) if 'balance' in col_map else 0.0
                            amount_val = debit_val if debit_val > 0 else credit_val
                            
                            txs.append({
                                'Date': date_str, 'Value Date': val_date_str, 'Description': desc_val,
                                'Ref / Cheque': ref_val, 'Amount': amount_val, 'Balance': balance_val,
                                'Debit': debit_val, 'Credit': credit_val
                            })
            if txs:
                return txs
    except Exception as e:
        print(f"Error in SBI dynamic table extraction: {e}")
            
    # 2. Fallback for line-based SBI statements
    txs = []
    anchor_indices = []
    for idx, line in enumerate(lines):
        # Match "DD/MM/YYYY DD/MM/YYYY" or "DD/MM/YY DD/MM/YY"
        match = re.match(r'^(\d{2}/\d{2}/\d{2,4})\s+(\d{2}/\d{2}/\d{2,4})\s+', line)
        if match:
            anchor_indices.append(idx)
            
    consumed = set()
    for i, anchor_idx in enumerate(anchor_indices):
        line = lines[anchor_idx]
        match = re.match(r'^(\d{2}/\d{2}/\d{2,4})\s+(\d{2}/\d{2}/\d{2,4})', line)
        date, val_date = match.groups()
        tokens = line.strip().split()
        
        balance = clean_amount(tokens[-1])
        
        def is_tok_amt(t):
            if not t or t == '-': return True
            tc = t.replace(",", "").strip()
            if re.search(r'[A-Za-z@/]', tc): return False
            try:
                float(tc)
                return True
            except ValueError:
                return False
                
        is_two_amounts = len(tokens) >= 5 and is_tok_amt(tokens[-2]) and is_tok_amt(tokens[-3])
        
        if is_two_amounts:
            debit_str = tokens[-3]
            credit_str = tokens[-2]
            debit = clean_amount(debit_str) if debit_str != '-' else 0.0
            credit = clean_amount(credit_str) if credit_str != '-' else 0.0
            amount = debit if debit > 0 else credit
            ref_no = ''
            anchor_desc = " ".join(tokens[2:-3])
        else:
            amount = clean_amount(tokens[-2])
            anchor_desc = " ".join(tokens[2:-2])
            ref_no = ''
            debit = 0.0
            credit = 0.0
            
        desc_parts = []
        
        # Check if preceding line is a header (like WDL TFR or DEP TFR) and not yet consumed
        if anchor_idx > 0 and (anchor_idx - 1) not in consumed:
            prev_line = lines[anchor_idx - 1]
            if is_sbi_header(prev_line):
                desc_parts.append(prev_line.strip())
                consumed.add(anchor_idx - 1)
                
        # Add anchor line's description part
        if anchor_desc.strip():
            desc_parts.append(anchor_desc.strip())
            
        consumed.add(anchor_idx)
        
        # Collect continuation lines up to the next transaction
        next_anchor_idx = anchor_indices[i+1] if i + 1 < len(anchor_indices) else len(lines)
        for idx in range(anchor_idx + 1, next_anchor_idx):
            if idx == next_anchor_idx - 1:
                # If next transaction has a preceding header, don't consume it here
                if is_sbi_header(lines[idx]):
                    break
                    
            curr_line = lines[idx].strip()
            if not curr_line:
                continue
            if is_noise_line(curr_line, "SBI"):
                continue
                
            desc_parts.append(curr_line)
            consumed.add(idx)
            
        desc = " ".join(desc_parts).strip()
        txs.append({
            'Date': date, 'Value Date': val_date, 'Description': desc,
            'Ref / Cheque': ref_no, 'Amount': amount, 'Balance': balance,
            'Debit': debit, 'Credit': credit
        })
        
    return txs

def parse_hdfc(pdf_file_obj, lines):
    txs = []
    
    # 1. Attempt PDFPlumber Page-by-Page Table Extraction (handles both HDFC2 grid format and HDFC3 multi-line column format)
    try:
        if hasattr(pdf_file_obj, "seek"):
            pdf_file_obj.seek(0)
        with pdfplumber.open(pdf_file_obj) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                if not tables:
                    continue
                for table in tables:
                    if not table or len(table) < 1:
                        continue
                    
                    # Case A: Multi-line newline joined columns per page (HDFC3 format)
                    has_multiline_date = False
                    for r in table:
                        if r and len(r) >= 5:
                            c0 = str(r[0] or "").strip()
                            if "\n" in c0 and len(re.findall(r'\b\d{2}[/\-\.]\d{2}[/\-\.]\d{2,4}\b', c0)) > 1:
                                has_multiline_date = True
                                break
                    
                    if has_multiline_date:
                        for row in table:
                            if not row or len(row) < 5:
                                continue
                            c0 = str(row[0] or "").strip()
                            if "\n" not in c0:
                                continue
                            
                            # Extract parallel lists
                            date_list = [d.strip() for d in c0.split('\n') if re.match(r'^\d{2}[/\-\.]\d{2}[/\-\.]\d{2,4}$', d.strip())]
                            if not date_list:
                                continue
                                
                            desc_lines = [d.strip() for d in str(row[1] or "").split('\n') if d.strip()]
                            ref_list = [r.strip() for r in str(row[2] or "").split('\n') if r.strip()] if len(row) > 2 else []
                            val_date_list = [v.strip() for v in str(row[3] or "").split('\n') if re.match(r'^\d{2}[/\-\.]\d{2}[/\-\.]\d{2,4}$', v.strip())] if len(row) > 3 else []
                            
                            # For amounts and closing balances:
                            if len(row) >= 7:
                                wdl_list = [clean_amount(w) for w in str(row[4] or "").split('\n') if w.strip() and clean_amount(w) > 0]
                                dep_list = [clean_amount(d) for d in str(row[5] or "").split('\n') if d.strip() and clean_amount(d) > 0]
                                bal_list = [clean_amount(b) for b in str(row[6] or "").split('\n') if b.strip() and clean_amount(b) > 0]
                            elif len(row) == 6:
                                wdl_list = [clean_amount(w) for w in str(row[3] or "").split('\n') if w.strip() and clean_amount(w) > 0]
                                dep_list = [clean_amount(d) for d in str(row[4] or "").split('\n') if d.strip() and clean_amount(d) > 0]
                                bal_list = [clean_amount(b) for b in str(row[5] or "").split('\n') if b.strip() and clean_amount(b) > 0]
                            else:
                                bal_list = []
                                wdl_list = []
                                dep_list = []
                                
                            # If we have exact balances for each date
                            if len(bal_list) == len(date_list):
                                page_text = page.extract_text() or ""
                                page_lines = [l.strip() for l in page_text.split('\n') if l.strip()]
                                
                                # Extract structured transaction blocks from page_lines
                                page_blocks = []
                                current_block = None
                                
                                for p_line in page_lines:
                                    if is_noise_line(p_line, "HDFC"):
                                        continue
                                    date_m = re.match(r'^(?:(\d+)\s+)?(\d{2}[/\-\.]\d{2}[/\-\.]\d{2,4})\s+(.+)', p_line)
                                    if date_m:
                                        if current_block:
                                            page_blocks.append(current_block)
                                        current_block = {
                                            'date': date_m.group(2),
                                            'raw_first_line': date_m.group(3),
                                            'extra_lines': []
                                        }
                                    elif current_block:
                                        current_block['extra_lines'].append(p_line)
                                        
                                if current_block:
                                    page_blocks.append(current_block)
                                    
                                # Compute description for each block
                                for b in page_blocks:
                                    f_tokens = b['raw_first_line'].split()
                                    t_end = len(f_tokens)
                                    while t_end > 0:
                                        tok_cand = f_tokens[t_end - 1]
                                        is_amt_tok = bool(re.match(r'^-?[\d,]+\.\d{2}$', tok_cand) and not re.search(r'[A-Za-z@]', tok_cand))
                                        is_dt_tok = bool(re.match(r'^\d{2}[/\-\.]\d{2}[/\-\.]\d{2,4}$', tok_cand))
                                        is_ref_tok = bool(tok_cand.isdigit() and len(tok_cand) >= 12)
                                        if is_amt_tok or is_dt_tok or is_ref_tok:
                                            t_end -= 1
                                        else:
                                            break
                                    b_desc = " ".join(f_tokens[:t_end]).strip()
                                    if b['extra_lines']:
                                        b_desc += " " + " ".join(b['extra_lines'])
                                    b['desc'] = b_desc.strip()
                                
                                first_credit = 0.0
                                first_debit = 0.0
                                first_amt = 0.0
                                if dep_list:
                                    first_credit = dep_list[0]
                                    first_amt = first_credit
                                elif wdl_list:
                                    first_debit = wdl_list[0]
                                    first_amt = first_debit
                                
                                used_block_indices = set()
                                for i, dt in enumerate(date_list):
                                    v_dt = val_date_list[i] if i < len(val_date_list) else dt
                                    r_no = ref_list[i] if i < len(ref_list) else ""
                                    bal = bal_list[i]
                                    
                                    # Match against page_blocks
                                    matched_desc = ""
                                    
                                    # 1. Direct index check
                                    if i < len(page_blocks) and i not in used_block_indices and page_blocks[i]['date'] == dt:
                                        matched_desc = page_blocks[i]['desc']
                                        used_block_indices.add(i)
                                    else:
                                        # 2. Search by date and ref/balance
                                        for b_idx, b in enumerate(page_blocks):
                                            if b_idx not in used_block_indices and b['date'] == dt:
                                                if (r_no and r_no in b['raw_first_line']) or f"{bal:,.2f}" in b['raw_first_line'] or f"{bal:.2f}" in b['raw_first_line']:
                                                    matched_desc = b['desc']
                                                    used_block_indices.add(b_idx)
                                                    break
                                        if not matched_desc:
                                            for b_idx, b in enumerate(page_blocks):
                                                if b_idx not in used_block_indices and b['date'] == dt:
                                                    matched_desc = b['desc']
                                                    used_block_indices.add(b_idx)
                                                    break
                                                    
                                    amt = first_amt if (len(txs) == 0 and i == 0) else 0.0
                                    cr = first_credit if (len(txs) == 0 and i == 0) else 0.0
                                    dr = first_debit if (len(txs) == 0 and i == 0) else 0.0
                                    
                                    txs.append({
                                        'Date': dt, 'Value Date': v_dt, 'Description': matched_desc,
                                        'Ref / Cheque': r_no, 'Amount': amt, 'Balance': bal,
                                        'Debit': dr, 'Credit': cr
                                    })
                        continue

                    # Case B: Standard row-by-row table (HDFC2 format)
                    for row in table:
                        if len(row) < 6:
                            continue
                        date_str = str(row[0] or "").strip()
                        if "\n" in date_str or not re.match(r'^\d{2}[/\-\.]\d{2}[/\-\.]\d{2,4}$', date_str):
                            continue
                        
                        desc = str(row[1] or "").strip().replace("\n", " ")
                        if len(row) >= 7:
                            ref_no = str(row[2] or "").strip()
                            val_date = str(row[3] or "").strip()
                            withdrawal = clean_amount(row[4])
                            deposit = clean_amount(row[5])
                            balance = clean_amount(row[6])
                        else:
                            ref_no = ""
                            val_date = str(row[2] or "").strip()
                            withdrawal = clean_amount(row[3])
                            deposit = clean_amount(row[4])
                            balance = clean_amount(row[5])
                            
                        amount = deposit if deposit > 0 else withdrawal
                        txs.append({
                            'Date': date_str, 'Value Date': val_date, 'Description': desc,
                            'Ref / Cheque': ref_no, 'Amount': amount, 'Balance': balance,
                            'Debit': withdrawal, 'Credit': deposit
                        })
        if txs:
            return txs
    except Exception as e:
        print(f"Error in HDFC table extraction: {e}")

    # 2. Comprehensive Line-by-Line Parsing (for HDFC1 and text statements)
    can_append = False
    date_regex = re.compile(r'^(?:(\d+)\s+)?(\d{2}[/\-\.]\d{2}[/\-\.]\d{2,4})\s+')
    
    for line in lines:
        line_clean = line.strip()
        if not line_clean or is_noise_line(line_clean, "HDFC"):
            can_append = False
            continue
            
        match = date_regex.match(line_clean)
        if match:
            date = match.group(2)
            remaining = line_clean[match.end():].strip()
            tokens = remaining.split()
            
            if len(tokens) >= 3:
                def is_amt(tok):
                    if not tok: return False
                    tok_c = tok.replace(",", "").strip()
                    tok_c = re.sub(r'(?i)(?:cr|dr|\(cr\)|\(dr\))$', '', tok_c)
                    if "/" in tok_c or (tok_c.startswith("-") and len(tok_c) <= 1): return False
                    try:
                        float(tok_c)
                        return True
                    except ValueError:
                        return False
                        
                balance = 0.0
                credit = 0.0
                debit = 0.0
                val_date = ""
                ref_no = ""
                desc = ""
                
                if is_amt(tokens[-1]):
                    balance = clean_amount(tokens[-1])
                    
                    if len(tokens) >= 3 and is_amt(tokens[-2]) and re.match(r'^\d{2}[/\-\.]\d{2}[/\-\.]\d{2,4}$', tokens[-3]):
                        amt = clean_amount(tokens[-2])
                        val_date = tokens[-3]
                        
                        if len(tokens) >= 4:
                            ref_no = tokens[-4]
                            desc = " ".join(tokens[:-4])
                        else:
                            desc = " ".join(tokens[:-3])
                        amount = amt
                    elif len(tokens) >= 4 and is_amt(tokens[-2]) and is_amt(tokens[-3]):
                        credit = clean_amount(tokens[-2])
                        debit = clean_amount(tokens[-3])
                        val_idx = -4
                        if len(tokens) >= 4 and re.match(r'^\d{2}[/\-\.]\d{2}[/\-\.]\d{2,4}$', tokens[val_idx]):
                            val_date = tokens[val_idx]
                            ref_idx = -5
                            if len(tokens) >= 5:
                                ref_no = tokens[ref_idx]
                                desc = " ".join(tokens[:ref_idx])
                            else:
                                desc = " ".join(tokens[:val_idx])
                        else:
                            desc = " ".join(tokens[:-3])
                        amount = credit if credit > 0 else debit
                    elif len(tokens) >= 2 and is_amt(tokens[-2]):
                        amt = clean_amount(tokens[-2])
                        desc = " ".join(tokens[:-2])
                        amount = amt
                    else:
                        desc = " ".join(tokens[:-1])
                        amount = 0.0
                        
                    txs.append({
                        'Date': date, 'Value Date': val_date, 'Description': desc,
                        'Ref / Cheque': ref_no, 'Amount': amount, 'Balance': balance,
                        'Debit': debit, 'Credit': credit
                    })
                    can_append = True
        elif txs and can_append:
            if not is_noise_line(line_clean, "HDFC") and not any(k in line_clean.lower() for k in (
                "opening balance", "closing balance", "brought forward", "carried forward", 
                "statement summary", "hdfc bank limited", "page no", "funds earmarked for hold",
                "registered office", "this is a computer generated statement"
            )):
                txs[-1]['Description'] += " " + line_clean
            else:
                can_append = False
                
    return txs

def parse_kotak(lines):
    txs = []
    can_append = False
    for line in lines:
        # Match "Index Date Month Year" pattern (e.g., "1 01 Apr 2025")
        match = re.match(r'^(\d+)\s+(\d{2}\s+[A-Za-z]{3}\s+\d{4})\s+', line)
        if match:
            idx, date = match.groups()
            tokens = line.strip().split()
            if len(tokens) >= 7:
                balance = clean_amount(tokens[-1])
                amount = clean_amount(tokens[-2])
                ref_no = tokens[-3]
                desc = " ".join(tokens[4:-3])
                
                txs.append({
                    'Date': date, 'Value Date': '', 'Description': desc,
                    'Ref / Cheque': ref_no, 'Amount': amount, 'Balance': balance,
                    'Debit': 0.0, 'Credit': 0.0
                })
                can_append = True
        elif txs and can_append:
            if is_noise_line(line, "KOTAK"):
                can_append = False
            else:
                txs[-1]['Description'] += " " + line.strip()
    return txs

def parse_idfc(lines):
    txs = []
    can_append = False
    for line in lines:
        # Match "DD-Mmm-YYYY" pattern (e.g., "03-Apr-2025")
        match = re.match(r'^(\d{2}-[A-Za-z]{3}-\d{4})\s+(\d{2}-[A-Za-z]{3}-\d{4})\s+', line)
        if match:
            date, val_date = match.groups()
            tokens = line.strip().split()
            if len(tokens) >= 4:
                balance = clean_amount(tokens[-1])
                amount = clean_amount(tokens[-2])
                # Check if third to last is a 6-digit cheque number
                ref_no = tokens[-3] if re.match(r'^\d{6}$', tokens[-3]) else ''
                desc_end_idx = -3 if ref_no else -2
                desc = " ".join(tokens[2:desc_end_idx])
                
                txs.append({
                    'Date': date, 'Value Date': val_date, 'Description': desc,
                    'Ref / Cheque': ref_no, 'Amount': amount, 'Balance': balance,
                    'Debit': 0.0, 'Credit': 0.0
                })
                can_append = True
        elif txs and can_append:
            if is_noise_line(line, "IDFC"):
                can_append = False
            else:
                txs[-1]['Description'] += " " + line.strip()
    return txs

def parse_rbl(pdf_file_obj, lines):
    txs = []
    use_layout_extraction = False
    try:
        if hasattr(pdf_file_obj, "seek"):
            pdf_file_obj.seek(0)
        with pdfplumber.open(pdf_file_obj) as pdf:
            if pdf.pages:
                use_layout_extraction = True
    except Exception as e:
        print(f"Error checking RBL layout extraction: {e}")
        
    if use_layout_extraction:
        try:
            if hasattr(pdf_file_obj, "seek"):
                pdf_file_obj.seek(0)
            with pdfplumber.open(pdf_file_obj) as pdf:
                for page in pdf.pages:
                    words = page.extract_words()
                    if not words:
                        continue
                    
                    # Group words by top coordinate
                    lines_dict = {}
                    for w in words:
                        found = False
                        for top_y in lines_dict:
                            if abs(w['top'] - top_y) < 3:
                                lines_dict[top_y].append(w)
                                found = True
                                break
                        if not found:
                            lines_dict[w['top']] = [w]
                            
                    # Sort lines by top coordinate
                    sorted_tops = sorted(lines_dict.keys())
                    
                    # Group lines into blocks based on vertical spacing
                    blocks = []
                    current_block = []
                    for i, top in enumerate(sorted_tops):
                        line_words = sorted(lines_dict[top], key=lambda x: x['x0'])
                        line_text = ' '.join([w['text'] for w in line_words]).strip()
                        
                        if i > 0:
                            diff = top - sorted_tops[i-1]
                            if diff > 15.0:  # Spacing threshold between transactions
                                if current_block:
                                    blocks.append(current_block)
                                current_block = []
                        current_block.append(line_text)
                    if current_block:
                        blocks.append(current_block)
                        
                    # Parse each block
                    for block in blocks:
                        anchor_idx = -1
                        anchor_tokens = None
                        for idx, line in enumerate(block):
                            tokens = line.split()
                            if len(tokens) >= 4:
                                if re.match(r'^\d{2}/\d{2}/\d{4}$', tokens[0]) and re.match(r'^\d{2}/\d{2}/\d{4}$', tokens[-3]):
                                    anchor_idx = idx
                                    anchor_tokens = tokens
                                    break
                                    
                        if anchor_idx != -1:
                            date_str = anchor_tokens[0]
                            val_date = anchor_tokens[-3]
                            amount = clean_amount(anchor_tokens[-2])
                            balance = clean_amount(anchor_tokens[-1])
                            
                            # Reconstruct description
                            desc_parts = []
                            for idx, line in enumerate(block):
                                if idx == anchor_idx:
                                    anchor_desc = " ".join(anchor_tokens[1:-3]).strip()
                                    if anchor_desc:
                                        desc_parts.append(anchor_desc)
                                else:
                                    desc_parts.append(line)
                            desc = " ".join(desc_parts).strip()
                            
                            txs.append({
                                'Date': date_str, 'Value Date': val_date, 'Description': desc,
                                'Ref / Cheque': '', 'Amount': amount, 'Balance': balance,
                                'Debit': 0.0, 'Credit': 0.0
                            })
            return txs
        except Exception as e:
            print(f"Error in RBL layout extraction, falling back: {e}")
            
    # Fallback to line-by-line parsing
    txs = []
    can_append = False
    for line in lines:
        match = re.match(r'^(\d{2}/\d{2}/\d{4})\s+', line)
        if match:
            date = match.group(1)
            tokens = line.strip().split()
            if len(tokens) >= 5:
                # Basic check to avoid footers in line-by-line fallback
                if re.match(r'^\d{2}/\d{2}/\d{4}$', tokens[-3]):
                    balance = clean_amount(tokens[-1])
                    amount = clean_amount(tokens[-2])
                    val_date = tokens[-3]
                    ref_no = tokens[-4] if re.match(r'^\d+$', tokens[-4]) else ''
                    desc_end_idx = -4 if ref_no else -3
                    desc = " ".join(tokens[1:desc_end_idx])
                    
                    txs.append({
                        'Date': date, 'Value Date': val_date, 'Description': desc,
                        'Ref / Cheque': ref_no, 'Amount': amount, 'Balance': balance,
                        'Debit': 0.0, 'Credit': 0.0
                    })
                    can_append = True
                else:
                    can_append = False
        elif txs and can_append:
            if is_noise_line(line, "RBL"):
                can_append = False
            else:
                txs[-1]['Description'] += " " + line.strip()
    return txs

def parse_modasa(lines):
    txs = []
    can_append = False
    for line in lines:
        # Match "DD/MM/YYYY" pattern
        match = re.match(r'^(\d{2}/\d{2}/\d{4})\s+', line)
        if match:
            date = match.group(1)
            tokens = line.strip().split()
            if len(tokens) >= 4:
                # Modasa bank has balance DR/CR suffix (e.g. 1796387.00 DR)
                has_suffix = tokens[-1] in ('DR', 'CR')
                balance_idx = -2 if has_suffix else -1
                balance = clean_amount(" ".join(tokens[balance_idx:]))
                amount = clean_amount(tokens[balance_idx - 1])
                
                # Check for numeric cheque/ref number
                ref_no = tokens[balance_idx - 2] if re.match(r'^\d+$', tokens[balance_idx - 2]) else ''
                desc_end_idx = balance_idx - 2 if ref_no else balance_idx - 1
                desc = " ".join(tokens[1:desc_end_idx])
                
                txs.append({
                    'Date': date, 'Value Date': '', 'Description': desc,
                    'Ref / Cheque': ref_no, 'Amount': amount, 'Balance': balance,
                    'Debit': 0.0, 'Credit': 0.0
                })
                can_append = True
        elif txs and can_append:
            if is_noise_line(line, "MODASA"):
                can_append = False
            else:
                txs[-1]['Description'] += " " + line.strip()
    return txs

def parse_boi(pdf_file_obj, lines):
    # 1. Dynamic Table Extraction (handles 7-col and multi-col BOI statements)
    try:
        if hasattr(pdf_file_obj, "seek"):
            pdf_file_obj.seek(0)
        with pdfplumber.open(pdf_file_obj) as pdf:
            txs = []
            col_map = {}
            header_col_count = 0
            
            for page in pdf.pages:
                tables = page.extract_tables()
                if not tables:
                    continue
                for table in tables:
                    if not table or len(table) < 1:
                        continue
                    
                    found_header = False
                    header_row = None
                    
                    for row_idx in range(min(5, len(table))):
                        row = [str(x or "").lower().replace("\n", " ").strip() for x in table[row_idx]]
                        has_date = any('date' in x or 'dt' in x for x in row)
                        has_desc = any('desc' in x or 'particular' in x or 'narration' in x for x in row)
                        
                        if has_date and has_desc:
                            header_row = table[row_idx]
                            col_map = {}
                            header_col_count = len(row)
                            for col_idx, col_name in enumerate(row):
                                col_name_clean = col_name.replace(".", " ").replace("(", " ").replace(")", " ").replace("/", " ").replace("-", " ").strip()
                                tokens = col_name_clean.split()
                                if 'date' in col_name_clean or 'dt' in col_name_clean:
                                    col_map['date'] = col_idx
                                elif 'desc' in col_name_clean or 'particular' in col_name_clean or 'narration' in col_name_clean:
                                    col_map['desc'] = col_idx
                                elif 'inst' in col_name_clean or 'chq' in col_name_clean or 'cheque' in col_name_clean or 'ref' in col_name_clean:
                                    col_map['ref'] = col_idx
                                elif 'sno' in col_name_clean or 'sl' in col_name_clean or 's no' in col_name_clean:
                                    col_map['sl_no'] = col_idx
                                elif 'debit' in col_name_clean or 'dr' in tokens:
                                    col_map['debit'] = col_idx
                                elif 'credit' in col_name_clean or 'cr' in tokens:
                                    col_map['credit'] = col_idx
                                elif 'balance' in col_name_clean or 'bal' in col_name_clean:
                                    col_map['balance'] = col_idx
                            found_header = True
                            break
                            
                    if not found_header and col_map and len(table[0]) == header_col_count:
                        start_row_idx = 0
                        parse_table = True
                    elif found_header and 'date' in col_map and 'desc' in col_map:
                        header_idx = table.index(header_row)
                        start_row_idx = header_idx + 1
                        parse_table = True
                    else:
                        parse_table = False
                        
                    if parse_table:
                        for r_idx in range(start_row_idx, len(table)):
                            row = table[r_idx]
                            if len(row) < header_col_count:
                                continue
                                
                            raw_date = str(row[col_map['date']] or "").strip()
                            date_m = re.search(r'\b\d{2}[/\-\.]\d{2}[/\-\.]\d{2,4}\b', raw_date)
                            
                            if not date_m:
                                # Continuation line
                                if txs and 'desc' in col_map and row[col_map['desc']]:
                                    desc_cont = str(row[col_map['desc']]).strip().replace("\n", " ")
                                    if desc_cont and not is_noise_line(desc_cont, "BOI") and not any(x in desc_cont.lower() for x in ("total", "carried forward", "brought forward", "page", "statement", "opening balance")):
                                        txs[-1]['Description'] += " " + desc_cont
                                continue
                                
                            date_str = date_m.group(0)
                            desc_val = str(row[col_map['desc']] or "").strip().replace("\n", " ")
                            
                            ref_val = ""
                            if 'ref' in col_map and row[col_map['ref']]:
                                raw_ref = str(row[col_map['ref']]).strip().replace("\n", " ")
                                ref_m = re.search(r'\b\d{5,22}\b', raw_ref)
                                if ref_m:
                                    ref_val = ref_m.group(0)
                                elif raw_ref and raw_ref != "-":
                                    ref_val = raw_ref
                                    
                            debit_val = clean_amount(row[col_map['debit']]) if 'debit' in col_map else 0.0
                            credit_val = clean_amount(row[col_map['credit']]) if 'credit' in col_map else 0.0
                            balance_val = clean_amount(row[col_map['balance']]) if 'balance' in col_map else 0.0
                            amount_val = debit_val if debit_val > 0 else credit_val
                            
                            txs.append({
                                'Date': date_str, 'Value Date': date_str, 'Description': desc_val,
                                'Ref / Cheque': ref_val, 'Amount': amount_val, 'Balance': balance_val,
                                'Debit': debit_val, 'Credit': credit_val
                            })
            if txs:
                return txs
    except Exception as e:
        print(f"Error in BOI dynamic table extraction: {e}")
        
    # 2. Line-by-line fallback for BOI
    txs = []
    can_append = False
    for line in lines:
        match = re.match(r'^(?:(\d+)\s+)?(\d{2}[/\-\.]\d{2}[/\-\.]\d{2,4})\s+', line.strip())
        if match:
            date = match.group(2)
            tokens = line.strip().split()
            if len(tokens) >= 4:
                has_suffix = tokens[-1].upper() in ('CR', 'DR', 'CR.', 'DR.')
                bal_idx = -2 if has_suffix else -1
                balance = clean_amount(tokens[bal_idx])
                
                # Check if both debit and credit tokens are present
                if len(tokens) >= 6 and re.match(r'^-?[\d,]+\.\d{2}$', tokens[bal_idx - 1]) and re.match(r'^-?[\d,]+\.\d{2}$', tokens[bal_idx - 2]):
                    credit = clean_amount(tokens[bal_idx - 1])
                    debit = clean_amount(tokens[bal_idx - 2])
                    amount = debit if debit > 0 else credit
                    desc_end = bal_idx - 2
                else:
                    amount = clean_amount(tokens[bal_idx - 1])
                    debit = 0.0
                    credit = 0.0
                    desc_end = bal_idx - 1
                    
                desc_start = 2 if match.group(1) else 1
                desc = " ".join(tokens[desc_start:desc_end])
                
                txs.append({
                    'Date': date, 'Value Date': date, 'Description': desc,
                    'Ref / Cheque': '', 'Amount': amount, 'Balance': balance,
                    'Debit': debit, 'Credit': credit
                })
                can_append = True
        elif txs and can_append:
            if is_noise_line(line, "BOI"):
                can_append = False
            else:
                txs[-1]['Description'] += " " + line.strip()
    return txs

def parse_sk(lines):
    txs = []
    can_append = False
    for line in lines:
        # Match "DD-MM-YYYY DD-MM-YYYY" pattern
        match = re.match(r'^(\d{2}-\d{2}-\d{4})\s+(\d{2}-\d{2}-\d{4})\s+', line)
        if match:
            date, val_date = match.groups()
            tokens = line.strip().split()
            if len(tokens) >= 6:
                balance = clean_amount(tokens[-1])
                amount = clean_amount(tokens[-2])
                ref_no = tokens[-3] if tokens[-3] != '-' else ''
                desc = " ".join(tokens[2:-3])
                
                txs.append({
                    'Date': date, 'Value Date': val_date, 'Description': desc,
                    'Ref / Cheque': ref_no, 'Amount': amount, 'Balance': balance,
                    'Debit': 0.0, 'Credit': 0.0
                })
                can_append = True
        elif txs and can_append:
            if is_noise_line(line, "SK"):
                can_append = False
            else:
                txs[-1]['Description'] += " " + line.strip()
    return txs

def parse_mehsana(lines):
    txs = []
    can_append = False
    for line in lines:
        # Match "DD-MMM-YYYY DD-MMM-YYYY" pattern
        match = re.match(r'^(\d{2}-[A-Za-z]{3}-\d{4})\s+(\d{2}-[A-Za-z]{3}-\d{4})\s+', line)
        if match:
            date, val_date = match.groups()
            tokens = line.strip().split()
            if len(tokens) >= 5:
                balance = clean_amount(tokens[-1])
                amount = clean_amount(tokens[-2])
                ref_no = tokens[-3] if re.match(r'^\d+$', tokens[-3]) else ''
                desc_end_idx = -3 if ref_no else -2
                desc = " ".join(tokens[2:desc_end_idx])
                
                txs.append({
                    'Date': date, 'Value Date': val_date, 'Description': desc,
                    'Ref / Cheque': ref_no, 'Amount': amount, 'Balance': balance,
                    'Debit': 0.0, 'Credit': 0.0
                })
                can_append = True
        elif txs and can_append:
            if is_noise_line(line, "MEHSANA"):
                can_append = False
            else:
                txs[-1]['Description'] += " " + line.strip()
    return txs

def parse_ubi1(lines):
    txs = []
    can_append = False
    for line in lines:
        # Match "Index DD/MM/YYYY" pattern
        match = re.match(r'^(\d+)\s+(\d{2}/\d{2}/\d{4})\s+', line)
        if match:
            idx, date = match.groups()
            tokens = line.strip().split()
            if len(tokens) >= 7:
                # Union Bank Format 1: amounts have suffix (Cr) or (Dr)
                # E.g., ['8485.00', '(Cr)', '213651.73', '(Cr)']
                balance = clean_amount(" ".join(tokens[-2:]))
                amount = clean_amount(" ".join(tokens[-4:-2]))
                tx_id = tokens[2]
                desc = " ".join(tokens[3:-4])
                
                txs.append({
                    'Date': date, 'Value Date': '', 'Description': desc,
                    'Ref / Cheque': tx_id, 'Amount': amount, 'Balance': balance,
                    'Debit': 0.0, 'Credit': 0.0
                })
                can_append = True
        elif txs and can_append:
            if is_noise_line(line, "UBI1"):
                can_append = False
            else:
                txs[-1]['Description'] += " " + line.strip()
    return txs

def parse_ubi2(lines):
    txs = []
    can_append = False
    for line in lines:
        # Match "Index DD-MM-YYYY" pattern
        match = re.match(r'^(\d+)\s+(\d{2}-\d{2}-\d{4})\s+', line)
        if match:
            idx, date = match.groups()
            tokens = line.strip().split()
            if len(tokens) >= 5:
                # UBI2 balance has Cr/Dr suffix (e.g. 1,10,703.06 Cr)
                has_suffix = tokens[-1] in ('Cr', 'Dr', 'CR', 'DR')
                balance_idx = -2 if has_suffix else -1
                balance = clean_amount(" ".join(tokens[balance_idx:]))
                amount = clean_amount(tokens[balance_idx - 1])
                
                # Check for cheque/ref number
                ref_no = tokens[balance_idx - 2] if re.match(r'^\d+$', tokens[balance_idx - 2]) else ''
                desc_end_idx = balance_idx - 2 if ref_no else balance_idx - 1
                desc = " ".join(tokens[2:desc_end_idx])
                
                txs.append({
                    'Date': date, 'Value Date': '', 'Description': desc,
                    'Ref / Cheque': ref_no, 'Amount': amount, 'Balance': balance,
                    'Debit': 0.0, 'Credit': 0.0
                })
                can_append = True
        elif txs and can_append:
            if is_noise_line(line, "UBI2"):
                can_append = False
            else:
                txs[-1]['Description'] += " " + line.strip()
    return txs

def clean_ubi3_amount_str(s):
    raw = re.sub(r'\([^\)]*\)', '', str(s)).strip()
    raw = raw.upper().replace("OO", "00").replace("O", "0").replace("Q", "0")
    raw = re.sub(r',(\d{2})$', r'.\1', raw)
    if '.' in raw:
        idx = raw.rfind('.')
        whole = re.sub(r'[^\d]', '', raw[:idx])
        dec = re.sub(r'[^\d]', '', raw[idx+1:])
        if len(whole) > 7:
            whole = whole[-7:]
        try:
            return float(f"{whole}.{dec[:2]}")
        except ValueError:
            return 0.0
    else:
        whole = re.sub(r'[^\d]', '', raw)
        if len(whole) > 7:
            whole = whole[-7:]
        try:
            return float(whole) if whole else 0.0
        except ValueError:
            return 0.0

def parse_ubi3_visual_row(date_str, text_line, prev_balance=None):
    amt_pattern = re.compile(r'\b([0-9OoQq]{1,8}(?:,[0-9OoQq]{2,3})*(?:[\.,]\s*[0-9OoQq]{2})?\s*\(\s*[DdCcOo0GgQq][Rr]\.?\s*\))')
    matches = list(amt_pattern.finditer(text_line))
    
    amt_val = 0.0
    bal_val = 0.0
    is_deb = True
    split_pos = -1
    
    if len(matches) >= 2:
        amt_match = matches[-2]
        bal_match = matches[-1]
        amt_str = amt_match.group(1)
        bal_str = bal_match.group(1)
        split_pos = amt_match.start()
        
        amt_val = clean_ubi3_amount_str(amt_str)
        bal_val = clean_ubi3_amount_str(bal_str)
            
        # Determine Debit vs Credit:
        if re.search(r'\(\s*[Dd][Rr]', amt_str, re.IGNORECASE):
            is_deb = True
        elif re.search(r'\(\s*[CcOo0GgQq][Rr]', amt_str, re.IGNORECASE):
            is_deb = False
        elif prev_balance is not None and prev_balance > 0:
            is_deb = (bal_val < prev_balance)
        else:
            is_deb = True
    else:
        # Fallback for transactions where numbers and (Cr)/(Dr) were on separate lines (e.g. 1170000,00 1339761.07 (Cr) (Cr))
        all_nums = list(re.finditer(r'\b([0-9OoQq]{1,8}(?:[.,][0-9OoQq]{2}))\b', text_line))
        if len(all_nums) >= 2:
            amt_match = all_nums[-2]
            bal_match = all_nums[-1]
            split_pos = amt_match.start()
            
            amt_val = clean_ubi3_amount_str(amt_match.group(1))
            bal_val = clean_ubi3_amount_str(bal_match.group(1))
            
            # Check context or compare with previous balance
            tags = re.findall(r'\(\s*([DdCcOo0GgQq][Rr])\.?\s*\)', text_line)
            if tags:
                first_tag = tags[0].upper()
                is_deb = first_tag.startswith('D')
            elif prev_balance is not None and prev_balance > 0:
                is_deb = (bal_val < prev_balance)
            else:
                is_deb = False
        else:
            return None
            
    deb_val = amt_val if is_deb else 0.0
    cred_val = amt_val if not is_deb else 0.0
    
    text_before = text_line[:split_pos].strip() if split_pos > 0 else text_line
    text_before = re.sub(r'^\d{2}-\d{2}-\s*', '', text_before)
    text_before = re.sub(r'^\d{2}-\d{2}-\d{4}\s*', '', text_before)
    text_before = re.sub(r'\b202[4567]\b', '', text_before).strip()
    
    tokens = text_before.split()
    tx_id = tokens[0] if tokens else ''
    desc = " ".join(tokens[1:]) if len(tokens) > 1 else ''
    
    return {
        'Date': date_str,
        'Value Date': '',
        'Description': desc,
        'Ref / Cheque': tx_id,
        'Amount': amt_val,
        'Balance': bal_val,
        'Debit': deb_val,
        'Credit': cred_val
    }

def parse_ubi3(pdf_file_obj, lines):
    txs = []
    
    if hasattr(pdf_file_obj, "seek"):
        pdf_file_obj.seek(0)
        
    try:
        with pdfplumber.open(pdf_file_obj) as pdf:
            last_known_balance = None
            for page_idx, page in enumerate(pdf.pages):
                words = page.extract_words(x_tolerance=3, y_tolerance=3)
                if not words:
                    continue
                    
                lines_dict = {}
                for w in words:
                    top_y = w['top']
                    found_line = False
                    for existing_top in lines_dict:
                        if abs(top_y - existing_top) < 4:
                            lines_dict[existing_top].append(w)
                            found_line = True
                            break
                    if not found_line:
                        lines_dict[top_y] = [w]
                        
                sorted_tops = sorted(lines_dict.keys())
                
                page_visual_lines = []
                for t in sorted_tops:
                    sorted_words = sorted(lines_dict[t], key=lambda x: x['x0'])
                    line_text = " ".join([w['text'] for w in sorted_words]).strip()
                    page_visual_lines.append((t, line_text))
                    
                p_len = len(page_visual_lines)
                v_idx = 0
                while v_idx < p_len:
                    t_y, l_str = page_visual_lines[v_idx]
                    
                    full_date_match = re.match(r'^(\d{2}-\d{2}-\d{4})', l_str)
                    split_date_match = re.match(r'^(\d{2}-\d{2}-)', l_str)
                    
                    if full_date_match:
                        date_str = full_date_match.group(1)
                        p = date_str.split('-')
                        if len(p) == 3 and p[1] in ('01', '02', '03') and p[2] == '2025':
                            date_str = f"{p[0]}-{p[1]}-2026"
                        tx = parse_ubi3_visual_row(date_str, l_str, prev_balance=last_known_balance)
                        if tx:
                            txs.append(tx)
                            if tx['Balance'] > 0:
                                last_known_balance = tx['Balance']
                        v_idx += 1
                        continue
                        
                    elif split_date_match:
                        date_prefix = split_date_match.group(1)
                        m_match = re.match(r'^(\d{2})-(\d{2})-', date_prefix)
                        day_str = m_match.group(1) if m_match else '01'
                        month_str = m_match.group(2) if m_match else '04'
                        year_str = "2026" if month_str in ('01', '02', '03') else "2025"
                        date_str = f"{day_str}-{month_str}-{year_str}"
                        
                        combined_text = l_str
                        consumed = 1
                        amt_p = re.compile(r'\b(\d{1,9}(?:,\d{2,3})*(?:\.\s*[0-9Oo]{1,2})?\s*\(\s*[DdCc][Rr]\.?\s*\))')
                        
                        for look in range(1, 6):
                            if v_idx + look < p_len:
                                next_t, next_l = page_visual_lines[v_idx + look]
                                if re.match(r'^\d{2}-\d{2}-', next_l):
                                    break
                                combined_text += " " + next_l
                                consumed = look + 1
                                if len(list(amt_p.finditer(combined_text))) >= 2:
                                    break
                                    
                        tx = parse_ubi3_visual_row(date_str, combined_text, prev_balance=last_known_balance)
                        if tx:
                            txs.append(tx)
                            if tx['Balance'] > 0:
                                last_known_balance = tx['Balance']
                        v_idx += consumed
                        continue
                        
                    v_idx += 1
                    
    except Exception as e:
        print(f"Error in coordinate-based UBI3 extraction: {e}")
        
    return txs

def parse_bob(pdf_file_obj, lines):
    raw_rows = []
    prefix_keywords = ('upi/', 'neft-', 'rtgs-', 'achdr/', 'sms ', 'by ', 'to ', 'chg/', 'interest', 'cash', 'imps', 'apy ', 'fis/', 'aeps/', 'by cash')
    
    if hasattr(pdf_file_obj, "seek"):
        pdf_file_obj.seek(0)
        
    with pdfplumber.open(pdf_file_obj) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            page_lines = text.split("\n")
            
            # Find table header
            header_idx = -1
            for idx, line in enumerate(page_lines):
                line_lower = line.lower()
                if "transaction" in line_lower and "description" in line_lower and "balance" in line_lower:
                    header_idx = idx
                    break
            
            if header_idx == -1:
                for idx, line in enumerate(page_lines):
                    line_lower = line.lower()
                    if "serial" in line_lower and "date" in line_lower and "debit" in line_lower:
                        header_idx = idx
                        break
                        
            if header_idx == -1:
                continue
                
            seen_first_date = False
            
            # Process lines after header
            for line in page_lines[header_idx + 1:]:
                line_str = line.strip()
                if not line_str:
                    continue
                    
                # Skip page footer noise
                if "discrepancy" in line_str.lower() or "computer-generated" in line_str.lower() or "page" in line_str.lower() or "note:" in line_str.lower():
                    break
                    
                match = re.match(r'^(\d+)\s+(\d{2}-\d{2}-\d{4})\s+(\d{2}-\d{2}-\d{4})\s+', line_str)
                if match:
                    seen_first_date = True
                    idx_str = match.group(1)
                    date = match.group(2)
                    val_date = match.group(3)
                    
                    remaining_str = line_str[match.end():].strip()
                    tokens = remaining_str.split()
                    
                    if len(tokens) >= 3:
                        balance = clean_amount(tokens[-1])
                        t_debit = tokens[-3]
                        t_credit = tokens[-2]
                        
                        debit = clean_amount(t_debit) if t_debit != '-' else 0.0
                        credit = clean_amount(t_credit) if t_credit != '-' else 0.0
                        amount = debit if debit > 0 else credit
                        
                        # Extra tokens before amount fields represent the description (and optional cheque number)
                        desc_tokens = tokens[:-3]
                        ref_no = ''
                        desc = ''
                        
                        if desc_tokens:
                            # If first token looks like a cheque number (e.g. 6-digit number)
                            if re.match(r'^\d{6}$', desc_tokens[0]):
                                ref_no = desc_tokens[0]
                                desc = " ".join(desc_tokens[1:])
                            else:
                                desc = " ".join(desc_tokens)
                        
                        raw_rows.append((True, {
                            'Date': date, 'Value Date': val_date, 'Description': desc,
                            'Ref / Cheque': ref_no, 'Amount': amount, 'Balance': balance,
                            'Debit': debit, 'Credit': credit
                        }))
                    else:
                        raw_rows.append((False, line_str))
                else:
                    # Before we see the first transaction date on this page,
                    # we ONLY keep lines that match our known prefix keywords.
                    if not seen_first_date:
                        is_prefix = line_str.lower().startswith(prefix_keywords)
                        if is_prefix:
                            raw_rows.append((False, line_str))
                    else:
                        raw_rows.append((False, line_str))
                    
        # Second pass: merge prefix and suffix description lines
        merged_txs = []
        prefix_buffer = []
        
        for i, (is_main, data) in enumerate(raw_rows):
            if is_main:
                tx = data
                if prefix_buffer:
                    if tx['Description']:
                        tx['Description'] = " ".join(prefix_buffer) + " " + tx['Description']
                    else:
                        tx['Description'] = " ".join(prefix_buffer)
                    prefix_buffer = []
                merged_txs.append(tx)
            else:
                line_str = data
                if is_noise_line(line_str, "BOB") or line_str.lower().startswith("opening balance"):
                    continue
                    
                is_prefix = line_str.lower().startswith(prefix_keywords)
                if is_prefix:
                    prefix_buffer.append(line_str)
                else:
                    if merged_txs:
                        if merged_txs[-1]['Description']:
                            merged_txs[-1]['Description'] += " " + line_str
                        else:
                            merged_txs[-1]['Description'] = line_str
                            
    return merged_txs

def parse_icici(pdf_file_obj, lines):
    use_table_extraction = False
    
    if hasattr(pdf_file_obj, "seek"):
        pdf_file_obj.seek(0)
        
    with pdfplumber.open(pdf_file_obj) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                if table and len(table[0]) >= 7:
                    col0 = (table[0][0] or "").lower().replace("\n", "")
                    if "si" in col0 or "sl" in col0 or "no" in col0 or "s." in col0:
                        use_table_extraction = True
                        break
            if use_table_extraction:
                break

    txs = []
    
    if use_table_extraction:
        try:
            if hasattr(pdf_file_obj, "seek"):
                pdf_file_obj.seek(0)
                
            with pdfplumber.open(pdf_file_obj) as pdf:
                col_map = {}
                header_col_count = 0
                
                for page in pdf.pages:
                    tables = page.extract_tables()
                    for table in tables:
                        if not table or len(table) < 1:
                            continue
                        
                        found_header = False
                        header_row = None
                        
                        for row_idx in range(min(5, len(table))):
                            row = [str(x or "").lower().replace("\n", " ").strip() for x in table[row_idx]]
                            has_date = any('date' in x for x in row)
                            has_remarks = any('remark' in x or 'description' in x or 'particular' in x for x in row)
                            
                            if has_date and has_remarks:
                                header_row = table[row_idx]
                                col_map = {}
                                header_col_count = len(row)
                                for col_idx, col_name in enumerate(row):
                                    col_name_clean = col_name.replace(".", " ").replace("(", " ").replace(")", " ").replace("/", " ").replace("-", " ").strip()
                                    if 'date' in col_name_clean:
                                        if 'value' in col_name_clean:
                                            col_map['val_date'] = col_idx
                                        elif 'posted' in col_name_clean:
                                            col_map['post_date'] = col_idx
                                        else:
                                            col_map['date'] = col_idx
                                    elif 'remark' in col_name_clean or 'description' in col_name_clean or 'particular' in col_name_clean:
                                        col_map['desc'] = col_idx
                                    elif 'cheque' in col_name_clean or 'chq' in col_name_clean or 'ref' in col_name_clean:
                                        col_map['ref'] = col_idx
                                    elif 'withdraw' in col_name_clean or 'debit' in col_name_clean or 'dr' in col_name_clean:
                                        col_map['debit'] = col_idx
                                    elif 'deposit' in col_name_clean or 'credit' in col_name_clean or 'cr' in col_name_clean:
                                        col_map['credit'] = col_idx
                                    elif 'balance' in col_name_clean or 'bal' in col_name_clean:
                                        col_map['balance'] = col_idx
                                    elif 'si' in col_name_clean or 'sl' in col_name_clean or 's' in col_name_clean:
                                        col_map['sl_no'] = col_idx
                                found_header = True
                                break
                                
                        if not found_header and col_map and len(table[0]) == header_col_count:
                            start_row_idx = 0
                            parse_table = True
                        elif found_header and 'date' in col_map and 'desc' in col_map:
                            header_idx = table.index(header_row)
                            start_row_idx = header_idx + 1
                            parse_table = True
                        else:
                            parse_table = False
                            
                        if parse_table:
                            for r_idx in range(start_row_idx, len(table)):
                                row = table[r_idx]
                                if len(row) < header_col_count:
                                    continue
                                
                                sl_no_val = ""
                                if 'sl_no' in col_map and row[col_map['sl_no']]:
                                    sl_no_val = str(row[col_map['sl_no']]).strip()
                                    
                                date_val = str(row[col_map['date']] or "").strip()
                                date_clean = date_val.replace("\n", " ").replace("\r", " ").strip()
                                
                                is_main = False
                                if re.search(r'\d', date_clean):
                                    is_main = True
                                elif sl_no_val and re.match(r'^\d+$', sl_no_val):
                                    is_main = True
                                    
                                if not is_main:
                                    # Continuation line
                                    if txs and 'desc' in col_map and row[col_map['desc']]:
                                        desc_cont = str(row[col_map['desc']]).strip().replace("\n", " ")
                                        if desc_cont and not is_noise_line(desc_cont, "ICICI") and not any(x in desc_cont.lower() for x in ("total", "carried forward", "brought forward", "page", "statement")):
                                            txs[-1]['Description'] += " " + desc_cont
                                    continue
                                    
                                desc_val = str(row[col_map['desc']] or "").strip().replace("\n", " ")
                                ref_val = str(row[col_map['ref']] or "").strip().replace("\n", " ") if 'ref' in col_map else ""
                                ref_match = re.search(r'\b\d{5,22}\b', ref_val)
                                if ref_match:
                                    ref_val = ref_match.group(0)
                                elif not ref_val or ref_val == "-":
                                    ref_val = ""
                                    
                                val_date_val = str(row[col_map['val_date']] or "").strip().replace("\n", " ") if 'val_date' in col_map else date_clean
                                
                                debit_val = clean_amount(row[col_map['debit']]) if 'debit' in col_map else 0.0
                                credit_val = clean_amount(row[col_map['credit']]) if 'credit' in col_map else 0.0
                                balance_val = clean_amount(row[col_map['balance']]) if 'balance' in col_map else 0.0
                                
                                amount_val = credit_val if credit_val > 0 else debit_val
                                
                                txs.append({
                                    'Date': date_clean, 'Value Date': val_date_val, 'Description': desc_val,
                                    'Ref / Cheque': ref_val, 'Amount': amount_val, 'Balance': balance_val,
                                    'Debit': debit_val, 'Credit': credit_val
                                })
        except Exception as e:
            print(f"Error in ICICI table parsing: {e}")
            
    else:
        # --- ICICI2 line-by-line parser ---
        raw_rows = []
        
        if hasattr(pdf_file_obj, "seek"):
            pdf_file_obj.seek(0)
            
        with pdfplumber.open(pdf_file_obj) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                page_lines = text.split("\n")
                
                # Find table header
                header_idx = -1
                for idx, line in enumerate(page_lines):
                    line_lower = line.lower()
                    if ("tran date" in line_lower and "value date" in line_lower and "particulars" in line_lower) or ("transaction date" in line_lower and "remarks" in line_lower):
                        header_idx = idx
                        break
                        
                if header_idx == -1:
                    continue
                    
                # Process lines after the header on this page
                for line in page_lines[header_idx + 1:]:
                    line_str = line.strip()
                    if not line_str:
                        continue
                        
                    # Stop processing this page once we hit the Total row
                    if line_str.lower().startswith(("total:", "total ")):
                        break
                        
                    # Match date format like DD-MM-YYYY or DD.MM.YYYY
                    match = re.match(r'^(?:\d+\s+)?(\d{2}[/\-.]\w{3,9}[/\-.]\d{2,4}|\d{2}[/\-.]\d{2}[/\-.]\d{2,4})\s+', line_str)
                    if match:
                        date = match.group(1)
                        tokens = line_str.split()
                        
                        # Find the index of the date token
                        date_token_idx = -1
                        for idx, tok in enumerate(tokens):
                            if date in tok or tok in date:
                                date_token_idx = idx
                                break
                        desc_start_idx = date_token_idx + 1 if date_token_idx != -1 else 1
                        
                        if len(tokens) >= desc_start_idx + 1:
                            has_indicator = tokens[-1].upper() in ('CR', 'DR')
                            last_idx = -2 if has_indicator else -1
                            
                            amounts = []
                            for idx in range(len(tokens) + last_idx, desc_start_idx - 1, -1):
                                tok = tokens[idx]
                                tok_clean = tok.replace(",", "").strip()
                                if re.match(r'^\d{1,2}[\-./](?:\d{2}|\w{3,9})[\-./]\d{2,4}$', tok) or re.match(r'^\d{6}$', tok_clean):
                                    break
                                if '.' not in tok_clean:
                                    break
                                try:
                                    val = float(tok_clean)
                                    amounts.insert(0, val)
                                except ValueError:
                                    break
                            
                            debit_val = 0.0
                            credit_val = 0.0
                            balance_val = 0.0
                            amount_val = 0.0
                            
                            num_amounts = len(amounts)
                            if num_amounts >= 3:
                                debit_val = amounts[-3]
                                credit_val = amounts[-2]
                                balance_val = amounts[-1]
                                amount_val = credit_val if credit_val > 0.0 else debit_val
                                desc_tokens = tokens[desc_start_idx:-3 - (1 if has_indicator else 0)]
                            elif num_amounts == 2:
                                amount_val = amounts[-2]
                                balance_val = amounts[-1]
                                desc_tokens = tokens[desc_start_idx:-2 - (1 if has_indicator else 0)]
                            elif num_amounts == 1:
                                amount_val = amounts[0]
                                desc_tokens = tokens[desc_start_idx:-1 - (1 if has_indicator else 0)]
                            else:
                                desc_tokens = tokens[desc_start_idx:]
                                
                            desc_val = " ".join(desc_tokens)
                            
                            val_date_val = ""
                            if len(desc_tokens) > 0 and re.match(r'^\d{2}[/\-.](?:\d{2}|\w{3,9})[/\-.]\d{2,4}$', desc_tokens[0]):
                                val_date_val = desc_tokens[0]
                                desc_val = " ".join(desc_tokens[1:])
                            
                            raw_rows.append((True, {
                                'Date': date, 'Value Date': val_date_val, 'Description': desc_val,
                                'Ref / Cheque': '', 'Amount': amount_val, 'Balance': balance_val,
                                'Debit': debit_val, 'Credit': credit_val
                            }))
                        else:
                            raw_rows.append((False, line_str))
                    else:
                        raw_rows.append((False, line_str))
        
        merged_txs = []
        prefix_buffer = []
        in_prefix_zone = False
        prefix_keywords = ('inf/', 'neft', 'rtgs', 'bil/', 'imps', 'chq', 'chg', 'cash', 'by ', 'to ', 'upi', 'paytm', 'phonepe', 'gpay', 'amazonpay', 'ybl', 'axl', 'ibld', 'transfer', 'clg', 'int.', 'interest')
        
        for is_main, data in raw_rows:
            if is_main:
                tx = data
                if prefix_buffer:
                    if tx['Description']:
                        tx['Description'] = " ".join(prefix_buffer) + " " + tx['Description']
                    else:
                        tx['Description'] = " ".join(prefix_buffer)
                    prefix_buffer = []
                merged_txs.append(tx)
                in_prefix_zone = False
            else:
                line_str = data
                if is_noise_line(line_str, "ICICI"):
                    continue
                    
                is_prefix = line_str.lower().startswith(prefix_keywords)
                if is_prefix:
                    in_prefix_zone = True
                    
                if in_prefix_zone:
                    prefix_buffer.append(line_str)
                else:
                    if merged_txs:
                        if merged_txs[-1]['Description']:
                            merged_txs[-1]['Description'] += " " + line_str
                        else:
                            merged_txs[-1]['Description'] = line_str
                            
        txs = merged_txs
        
    return txs

def parse_axis(pdf_file_obj):
    txs = []
    
    if hasattr(pdf_file_obj, "seek"):
        pdf_file_obj.seek(0)
        
    with pdfplumber.open(pdf_file_obj) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                if table and len(table[0]) >= 7:
                    for row in table[1:]:
                        if not row or len(row) < 7:
                            continue
                        
                        s_no = (row[0] or "").strip()
                        date = (row[1] or "").strip().replace("\n", " ").replace("\r", " ")
                        
                        if re.match(r'^\d+$', s_no) and re.match(r'^\d{2}/\d{2}/\d{4}', date):
                            val_date = (row[2] or "").strip().replace("\n", " ").replace("\r", " ")
                            desc = (row[3] or "").strip().replace("\n", " ").replace("\r", " ")
                            
                            amount = clean_amount(row[4])
                            drcr = (row[5] or "").strip().upper()
                            balance = clean_amount(row[6])
                            ref_no = (row[7] or "").strip().replace("\n", " ").replace("\r", " ") if len(row) > 7 else ''
                            
                            debit = amount if drcr == 'DR' else 0.0
                            credit = amount if drcr == 'CR' else 0.0
                            
                            txs.append({
                                'Date': date, 'Value Date': val_date, 'Description': desc,
                                'Ref / Cheque': ref_no, 'Amount': amount, 'Balance': balance,
                                'Debit': debit, 'Credit': credit
                            })
                            
    return txs

def parse_prime(pdf_file_obj):
    txs = []
    
    if hasattr(pdf_file_obj, "seek"):
        pdf_file_obj.seek(0)
        
    with pdfplumber.open(pdf_file_obj) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                if not table or len(table[0]) < 5:
                    continue
                    
                # Identify the transaction table by looking for narration + withdrawal/deposit headers
                headers = [str(h or '').lower().replace("\n", " ") for h in table[0]]
                has_narration = any("narration" in h or "particulars" in h for h in headers)
                has_amounts = any("withdrawal" in h or "deposit" in h or "debit" in h or "credit" in h for h in headers)
                if not (has_narration and has_amounts):
                    continue
                    
                for row in table[1:]:
                    if not row or len(row) < 5:
                        continue
                        
                    date_cell = (row[0] or "").strip()
                    if not date_cell:
                        continue
                        
                    desc = (row[1] or "").strip().replace("\n", " ").replace("\r", " ")
                    desc_lower = desc.lower()
                    
                    # Skip opening/closing balance and summary rows
                    if any(kw in desc_lower for kw in ("opening balance", "closing balance", "brought forward", "carried forward", "total", "transaction count")):
                        continue
                        
                    # Date cell may contain two lines: "DD/MM/YYYY\nDD/MM/YYYY" (txn date + value date)
                    date_parts = date_cell.split("\n")
                    date = date_parts[0].strip()
                    val_date = date_parts[1].strip() if len(date_parts) > 1 else date
                    
                    # Validate date format
                    if not re.match(r'^\d{2}/\d{2}/\d{4}', date):
                        continue
                        
                    ref_no = (row[2] or "").strip().replace("\n", " ").replace("\r", " ")
                    withdrawal = clean_amount(row[3]) if len(row) > 3 else 0.0
                    deposit = clean_amount(row[4]) if len(row) > 4 else 0.0
                    balance = clean_amount(row[5]) if len(row) > 5 else 0.0
                    
                    amount = deposit if deposit > 0 else withdrawal
                    if amount == 0.0:
                        continue
                        
                    txs.append({
                        'Date': date, 'Value Date': val_date, 'Description': desc,
                        'Ref / Cheque': ref_no, 'Amount': amount, 'Balance': balance,
                        'Debit': withdrawal, 'Credit': deposit
                    })
                    
    return txs

# Helper function to guess the Debit/Credit direction of the very first transaction row in the statement
def guess_first_txn_direction(tx, bank_type):
    desc = tx['Description'].lower()
    bal = tx['Balance']
    amt = abs(tx['Amount'])
    
    # 1. Standard transaction indicators check
    credits = ('payment from', 'received from', 'by transfer', 'credit', 'deposit', 'dep ', 'dep:', 'interest', 'cr', 'refund', 'received', 'by clg')
    debits = ('payment to', 'paid to', 'to transfer', 'debit', 'withdrawal', 'wdl', 'charges', 'rent', 'dr', 'fee', 'to clg', 'payment')
    
    if any(k in desc for k in credits):
        return True # Credit
    if any(k in desc for k in debits):
        return False # Debit
        
    # 2. Bank-specific narration indicator mappings
    if bank_type == "BOI":
        if 'cr/' in desc or '/cr/' in desc or 'cr ' in desc:
            return True
        return False
    elif bank_type == "HDFC":
        if '- cr -' in desc or ' cr ' in desc:
            return True
        return False
    elif bank_type == "IDFC":
        if '/cr/' in desc or 'interest' in desc:
            return True
        return False
    elif bank_type == "KOTAK":
        if 'rev-' in desc or 'deposit' in desc or 'cr' in desc:
            return True
        return False
    elif bank_type == "RBL":
        if 'ach cr' in desc or 'upi/cr' in desc or ' cr ' in desc:
            return True
        return False
        
    # 3. Positive opening balance heuristic
    # If Credit yields a negative opening balance, but Debit yields a positive one, assume Debit.
    if bal - amt < 0.0 and bal + amt >= 0.0:
        return False # Debit
        
    return True # Default fallback

# Resolve Debits/Credits using running balance progression
def resolve_debits_credits(txs, bank_type, is_reverse=False):
    if not txs:
        return txs
        
    # If any transaction already has populated Debit or Credit (ICICI1/BOB/AXIS/PRIME/SBI2/BOI format), skip resolution
    if bank_type in ("ICICI", "BOB", "AXIS", "PRIME", "SBI", "BOI", "GENERIC", "UBI3") and any(tx.get('Debit', 0.0) > 0.0 or tx.get('Credit', 0.0) > 0.0 for tx in txs):
        return txs
        
    # If the statement is reverse chronological (newest first), reverse it for progression calculation
    if is_reverse:
        txs.reverse()
        
    prev_bal = None
    for i, tx in enumerate(txs):
        bal = tx['Balance']
        amt = abs(tx['Amount'])
        
        if prev_bal is not None:
            diff = bal - prev_bal
            # Account for rounding tolerance
            if diff > 0.02:
                tx['Credit'] = diff
                tx['Debit'] = 0.0
            elif diff < -0.02:
                tx['Debit'] = abs(diff)
                tx['Credit'] = 0.0
            else:
                # No change in balance
                tx['Debit'] = 0.0
                tx['Credit'] = 0.0
        else:
            # First transaction row: if already populated (e.g. from table), preserve it; otherwise guess
            if tx.get('Credit', 0.0) == 0.0 and tx.get('Debit', 0.0) == 0.0:
                is_credit = guess_first_txn_direction(tx, bank_type)
                if is_credit:
                    tx['Credit'] = amt
                    tx['Debit'] = 0.0
                else:
                    tx['Debit'] = amt
                    tx['Credit'] = 0.0
            else:
                if tx.get('Credit', 0.0) > 0:
                    tx['Amount'] = tx['Credit']
                elif tx.get('Debit', 0.0) > 0:
                    tx['Amount'] = tx['Debit']
                
        prev_bal = bal
        
    # Reverse it back to match the original PDF page listing order
    if is_reverse:
        txs.reverse()
        
    return txs

# Format worksheet layout and auto-adjust widths
def format_worksheet(ws):
    column_order = [
        'Date',
        'Narration',
        'Debit (Rs)',
        'Credit (Rs)',
        'Balance (Rs)',
        'Chq / Ref No',
        'Data Source',
        'Source File'
    ]
    
    # Style definitions
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid") # Corporate Steel Blue
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=11)
    
    alt_row_fill = PatternFill(start_color="F2F6FA", end_color="F2F6FA", fill_type="solid") # Soft light blue-gray
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    
    left_align = Alignment(horizontal="left", vertical="center")
    center_align = Alignment(horizontal="center", vertical="center")
    right_align = Alignment(horizontal="right", vertical="center")
    
    currency_format = '#,##0.00'
    
    # Format Headers
    for col_num in range(1, len(column_order) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border
        
    # Format Data cells
    for row_num in range(2, ws.max_row + 1):
        for col_num, col_name in enumerate(column_order, 1):
            cell = ws.cell(row=row_num, column=col_num)
            cell.font = data_font
            cell.border = thin_border
            if row_num % 2 == 0:
                cell.fill = alt_row_fill
            
            # Format number columns
            if col_name in ('Debit (Rs)', 'Credit (Rs)', 'Balance (Rs)'):
                cell.number_format = currency_format
                cell.alignment = right_align
                if cell.value is not None:
                    try:
                        cell.value = float(cell.value)
                    except ValueError:
                        pass
            elif col_name in ('Date', 'Data Source'):
                cell.alignment = center_align
            else:
                cell.alignment = left_align
                
    # Auto-adjust column width
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            val_str = str(cell.value or '')
            if cell.row == 1:
                val_str += '   ' # Add padding for filter dropdown arrows in Excel
            max_len = max(max_len, len(val_str))
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 48)

    # Format as Excel Table (without default row stripes)
    if ws.max_row > 1:
        sheet_name_clean = re.sub(r'[^A-Za-z0-9]', '', ws.title)
        table_name = f"Table_{sheet_name_clean}"
        max_col_letter = get_column_letter(len(column_order))
        ref_range = f"A1:{max_col_letter}{ws.max_row}"
        tab = Table(displayName=table_name, ref=ref_range)
        style = TableStyleInfo(
            name="TableStyleLight9",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=False, # Use our custom alt row fill instead
            showColumnStripes=False
        )
        tab.tableStyleInfo = style
        try:
            ws.add_table(tab)
        except Exception:
            pass

# Format Tally Voucher worksheet layout
def format_tally_worksheet(ws):
    column_order = [
        'Voucher Date',
        'Voucher Type Name',
        'Voucher Number',
        'Ledger Name',
        'Ledger Amount',
        'Voucher Narration',
        'Ledger Amount Dr/Cr'
    ]
    
    # Style definitions
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid") # Corporate Steel Blue
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=11)
    
    alt_row_fill = PatternFill(start_color="F2F6FA", end_color="F2F6FA", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    
    left_align = Alignment(horizontal="left", vertical="center")
    center_align = Alignment(horizontal="center", vertical="center")
    right_align = Alignment(horizontal="right", vertical="center")
    
    currency_format = '#,##0.00'
    
    # Format Headers
    for col_num in range(1, len(column_order) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border
        
    # Format Data cells
    for row_num in range(2, ws.max_row + 1):
        for col_num, col_name in enumerate(column_order, 1):
            cell = ws.cell(row=row_num, column=col_num)
            cell.font = data_font
            cell.border = thin_border
            if row_num % 2 == 0:
                cell.fill = alt_row_fill
            
            # Format number columns
            if col_name == 'Ledger Amount':
                cell.number_format = currency_format
                cell.alignment = right_align
                if cell.value is not None:
                    try:
                        cell.value = float(cell.value)
                    except ValueError:
                        pass
            elif col_name in ('Voucher Date', 'Voucher Type Name', 'Voucher Number', 'Ledger Amount Dr/Cr'):
                cell.alignment = center_align
            else:
                cell.alignment = left_align
                
    # Auto-adjust column width
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            val_str = str(cell.value or '')
            if cell.row == 1:
                val_str += '   ' # Add padding for filter dropdown arrows in Excel
            max_len = max(max_len, len(val_str))
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 48)

    # Format as Excel Table (without default row stripes)
    if ws.max_row > 1:
        sheet_name_clean = re.sub(r'[^A-Za-z0-9]', '', ws.title)
        table_name = f"Table_{sheet_name_clean}"
        max_col_letter = get_column_letter(len(column_order))
        ref_range = f"A1:{max_col_letter}{ws.max_row}"
        tab = Table(displayName=table_name, ref=ref_range)
        style = TableStyleInfo(
            name="TableStyleLight9",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=False, # Use our custom alt row fill instead
            showColumnStripes=False
        )
        tab.tableStyleInfo = style
        try:
            ws.add_table(tab)
        except Exception:
            pass

# Format Summary worksheet layout
def format_summary_worksheet(ws):
    column_order = [
        'Source File',
        'Bank',
        'Opening Balance',
        'Total Credits',
        'Total Debits',
        'Closing Balance'
    ]
    
    # Style definitions
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid") # Corporate Steel Blue
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=11)
    bold_font = Font(name="Calibri", size=11, bold=True)
    
    alt_row_fill = PatternFill(start_color="F2F6FA", end_color="F2F6FA", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    
    left_align = Alignment(horizontal="left", vertical="center")
    center_align = Alignment(horizontal="center", vertical="center")
    right_align = Alignment(horizontal="right", vertical="center")
    
    currency_format = '#,##0.00'
    
    # Format Headers
    for col_num in range(1, len(column_order) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border
        
    # Format Data cells
    for row_num in range(2, ws.max_row + 1):
        is_total_row = (ws.cell(row=row_num, column=1).value == 'Total')
        for col_num, col_name in enumerate(column_order, 1):
            cell = ws.cell(row=row_num, column=col_num)
            cell.font = bold_font if is_total_row else data_font
            cell.border = thin_border
            if not is_total_row and row_num % 2 == 0:
                cell.fill = alt_row_fill
            
            # Format number columns
            if col_name in ('Opening Balance', 'Total Credits', 'Total Debits', 'Closing Balance'):
                cell.number_format = currency_format
                cell.alignment = right_align
                if cell.value is not None:
                    try:
                        cell.value = float(cell.value)
                    except ValueError:
                        pass
            elif col_name == 'Bank':
                cell.alignment = center_align
            else:
                cell.alignment = left_align
                
    # Auto-adjust column width
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            val_str = str(cell.value or '')
            if cell.row == 1:
                val_str += '   ' # Add padding for filter dropdown arrows in Excel
            max_len = max(max_len, len(val_str))
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 48)

    # Format as Excel Table (without default row stripes)
    if ws.max_row > 1:
        sheet_name_clean = re.sub(r'[^A-Za-z0-9]', '', ws.title)
        table_name = f"Table_{sheet_name_clean}"
        max_col_letter = get_column_letter(len(column_order))
        ref_range = f"A1:{max_col_letter}{ws.max_row}"
        tab = Table(displayName=table_name, ref=ref_range)
        style = TableStyleInfo(
            name="TableStyleLight9",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=False, # Use our custom alt row fill instead
            showColumnStripes=False
        )
        tab.tableStyleInfo = style
        try:
            ws.add_table(tab)
        except Exception:
            pass

# Transform transactions into double-entry Tally Voucher rows
def generate_tally_voucher_df(df_sorted):
    tally_rows = []
    records = df_sorted.to_dict('records')
    for i, row in enumerate(records, 1):
        date = row.get('Txn Date', '')
        narration = row.get('Description', '')
        
        # Safe float conversion
        debit = 0.0
        if row.get('Debit / Withdrawal') is not None and str(row.get('Debit / Withdrawal')).strip():
            try:
                debit = float(str(row.get('Debit / Withdrawal')).replace(',', ''))
            except ValueError:
                pass
                
        credit = 0.0
        if row.get('Credit / Deposit') is not None and str(row.get('Credit / Deposit')).strip():
            try:
                credit = float(str(row.get('Credit / Deposit')).replace(',', ''))
            except ValueError:
                pass
                
        if debit > 0:
            vtype = "Payment"
            amount = debit
            suspense_drcr = "Dr"
            bank_drcr = "Cr"
        else:
            vtype = "Receipt"
            amount = credit
            suspense_drcr = "Cr"
            bank_drcr = "Dr"
            
        # Row 1: Suspense
        tally_rows.append({
            'Voucher Date': date,
            'Voucher Type Name': vtype,
            'Voucher Number': i,
            'Ledger Name': 'Suspense',
            'Ledger Amount': amount,
            'Voucher Narration': narration,
            'Ledger Amount Dr/Cr': suspense_drcr
        })
        
        # Row 2: Bank Account
        tally_rows.append({
            'Voucher Date': date,
            'Voucher Type Name': vtype,
            'Voucher Number': i,
            'Ledger Name': row.get('Bank') or 'Bank Account',
            'Ledger Amount': amount,
            'Voucher Narration': narration,
            'Ledger Amount Dr/Cr': bank_drcr
        })
        
    return pd.DataFrame(tally_rows)

def parse_generic_blocks(lines):
    """
    Fallback parser for completely non-tabular PDF layouts (e.g. text blocks, key-values, 
    or multi-line paragraph flows).
    Groups lines into block-logical transactions starting from each Date.
    Finds dates, decimal numbers (amounts), direction labels, and narative descriptions.
    """
    print(f"[Parser] Heuristic Block Parser: Scanning {len(lines)} lines for non-tabular layouts...")
    txs = []
    i = 0
    n_lines = len(lines)
    
    # Common date patterns (DD/MM/YYYY, DD-Mmm-YYYY, YYYY-MM-DD, etc.)
    date_pattern = r'\b(\d{2}[/\-.]\d{2}[/\-.]\d{2,4}|\d{2}[\s\-./][A-Za-z]{3,10}[\s\-./]\d{2,4}|\d{4}[\-./]\d{2}[\-./]\d{2})\b'
    
    while i < n_lines:
        line_str = lines[i].strip()
        if not line_str:
            i += 1
            continue
            
        # Check if this line starts a new transaction block (contains a date)
        date_match = re.search(date_pattern, line_str)
        if date_match:
            date_val = date_match.group(1)
            
            # Form a block of lines until the next date or max 5 lines
            block_lines = [line_str]
            j = i + 1
            while j < n_lines:
                next_line = lines[j].strip()
                if not next_line:
                    j += 1
                    continue
                # If we hit another date, stop the block
                if re.search(date_pattern, next_line):
                    break
                block_lines.append(next_line)
                j += 1
                if len(block_lines) >= 6:
                    break
            
            # Parse the block_lines
            amounts_found = []
            desc_parts = []
            debit_val = 0.0
            credit_val = 0.0
            balance_val = 0.0
            
            has_dr_label = False
            has_cr_label = False
            
            for b_line in block_lines:
                b_line_lower = b_line.lower()
                
                # Check for DR/CR indicators in this line
                if 'dr' in b_line_lower or 'debit' in b_line_lower or 'withdraw' in b_line_lower:
                    has_dr_label = True
                if 'cr' in b_line_lower or 'credit' in b_line_lower or 'deposit' in b_line_lower:
                    has_cr_label = True
                
                # Find all potential decimal numbers in this line
                tokens = b_line.split()
                for tok in tokens:
                    tok_clean = tok.replace(",", "").strip()
                    # Strip trailing Cr/Dr if attached
                    tok_clean = re.sub(r'(?i)(?:cr|dr|clg|chq)$', '', tok_clean)
                    
                    if re.match(r'^\d+\.\d{2}$', tok_clean) or (re.match(r'^\d+$', tok_clean) and len(tok_clean) > 1):
                        try:
                            val = float(tok_clean)
                            if 2000 <= val <= 2030 and val == int(val):
                                continue
                            if val < 10.0 and val == int(val):
                                continue
                            if '.' not in tok_clean and val > 99999:
                                continue
                            
                            amounts_found.append((tok, val, b_line_lower))
                        except ValueError:
                            pass
                
                # Extract description
                clean_desc_line = b_line
                clean_desc_line = clean_desc_line.replace(date_val, "")
                for amt_str, _, _ in amounts_found:
                    clean_desc_line = clean_desc_line.replace(amt_str, "")
                
                clean_desc_line = clean_desc_line.strip()
                if clean_desc_line and not is_noise_line(clean_desc_line, "GENERIC"):
                    desc_parts.append(clean_desc_line)
            
            # Resolve amounts
            vals = [x[1] for x in amounts_found]
            
            if len(vals) >= 2:
                amount_val = vals[0]
                balance_val = vals[-1]
                
                first_amt_line = amounts_found[0][2]
                if 'dr' in first_amt_line or 'debit' in first_amt_line or 'withdraw' in first_amt_line:
                    debit_val = amount_val
                elif 'cr' in first_amt_line or 'credit' in first_amt_line or 'deposit' in first_amt_line:
                    credit_val = amount_val
                elif has_dr_label:
                    debit_val = amount_val
                elif has_cr_label:
                    credit_val = amount_val
                else:
                    debit_val = amount_val
            elif len(vals) == 1:
                amount_val = vals[0]
                balance_val = 0.0
                
                first_amt_line = amounts_found[0][2]
                if 'dr' in first_amt_line or 'debit' in first_amt_line or 'withdraw' in first_amt_line:
                    debit_val = amount_val
                elif 'cr' in first_amt_line or 'credit' in first_amt_line or 'deposit' in first_amt_line:
                    credit_val = amount_val
                elif has_dr_label:
                    debit_val = amount_val
                elif has_cr_label:
                    credit_val = amount_val
                else:
                    debit_val = amount_val
            else:
                amount_val = 0.0
                balance_val = 0.0
                
            desc_val = " ".join(desc_parts).strip()
            desc_val = re.sub(r'\s+', ' ', desc_val)
            desc_val = re.sub(r'^[\-\s./]+|[\-\s./]+$', '', desc_val)
            
            if amount_val > 0.0 or desc_val:
                txs.append({
                    'Date': date_val, 'Value Date': date_val, 'Description': desc_val,
                    'Ref / Cheque': '', 'Amount': amount_val, 'Balance': balance_val,
                    'Debit': debit_val, 'Credit': credit_val
                })
            
            i = j - 1
            
        i += 1
        
    return txs

def parse_generic(pdf_file_obj, lines):
    try:
        with open("debug_generic_parser.txt", "w", encoding="utf-8") as f_dbg:
            f_dbg.write("--- GENERIC PARSER DIAGNOSTICS START ---\n")
    except Exception:
        pass
        
    # Try table extraction first
    try:
        if hasattr(pdf_file_obj, "seek"):
            pdf_file_obj.seek(0)
        with pdfplumber.open(pdf_file_obj) as pdf:
            txs = []
            col_map = {}
            header_col_count = 0
            
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if not table or len(table) < 1:
                        continue
                    
                    header_row = None
                    found_header = False
                    
                    for row_idx in range(min(5, len(table))):
                        row = [str(x or "").lower().strip() for x in table[row_idx]]
                        has_date_hdr = any('date' in x or 'dt' in x for x in row)
                        has_desc_hdr = any('desc' in x or 'particular' in x or 'narration' in x or 'remark' in x or 'detail' in x or 'description' in x for x in row)
                        
                        if has_date_hdr and has_desc_hdr:
                            header_row = table[row_idx]
                            col_map = {}
                            header_col_count = len(row)
                            for col_idx, col_name in enumerate(row):
                                col_name_clean = col_name.replace(".", " ").replace("(", " ").replace(")", " ").replace("/", " ").replace("-", " ").strip()
                                tokens = col_name_clean.split()
                                
                                if 'date' in col_name or 'dt' in col_name:
                                    if 'val' in col_name or 'value' in col_name:
                                        col_map['val_date'] = col_idx
                                    else:
                                        col_map['date'] = col_idx
                                elif 'desc' in col_name or 'particular' in col_name or 'narration' in col_name or 'remark' in col_name or 'detail' in col_name or 'description' in col_name:
                                    col_map['desc'] = col_idx
                                elif 'chq' in col_name or 'cheque' in col_name or 'ref' in col_name or 'instrument' in col_name or 'doc' in col_name:
                                    col_map['ref'] = col_idx
                                elif any(x in col_name_clean for x in ('debit', 'withdrawal', 'withdraw', 'payment', 'outflow', 'debited', 'wdl')) or 'dr' in tokens or 'out' in tokens or 'w/d' in tokens:
                                    col_map['debit'] = col_idx
                                elif any(x in col_name_clean for x in ('credit', 'deposit', 'receipt', 'inflow', 'credited', 'dep')) or 'cr' in tokens or 'in' in tokens:
                                    col_map['credit'] = col_idx
                                elif 'balance' in col_name or 'bal' in col_name:
                                    col_map['balance'] = col_idx
                                elif 'amount' in col_name or 'amt' in col_name:
                                    col_map['amount'] = col_idx
                                elif 'type' in col_name or 'cr/dr' in col_name or 'drcr' in col_name:
                                    col_map['type'] = col_idx
                            found_header = True
                            break
                            
                    if not found_header and col_map and len(table[0]) == header_col_count:
                        # Re-use previous page's col_map, parsing from row 0
                        start_row_idx = 0
                        parse_table = True
                    elif found_header and 'date' in col_map and 'desc' in col_map:
                        header_idx = table.index(header_row)
                        start_row_idx = header_idx + 1
                        parse_table = True
                    else:
                        parse_table = False
                        
                    if parse_table:
                        for r_idx in range(start_row_idx, len(table)):
                            row = table[r_idx]
                            if len(row) < header_col_count:
                                continue
                            
                            # Resolve date_val, supporting multi-line wrapped dates
                            raw_date_cell = str(row[col_map['date']] or "").strip()
                            date_lines = [l.strip() for l in raw_date_cell.split('\n') if l.strip()]
                            date_val = ""
                            if date_lines:
                                first_line = date_lines[0]
                                date_patterns = [
                                    r'\b\d{2}[/\-.]\d{2}[/\-.]\d{2,4}\b',
                                    r'\b\d{2}[\s\-./][A-Za-z]{3}[\s\-./]\d{2,4}\b',
                                    r'\b\d{4}[\-./]\d{2}[\-./]\d{2}\b'
                                ]
                                if any(re.search(pat, first_line) for pat in date_patterns):
                                    date_val = first_line
                                else:
                                    joined = " ".join(date_lines)
                                    date_val = re.sub(r'([\-./])\s+', r'\1', joined)
                                    
                            # Check if the date contains any digits
                            if not re.search(r'\d', date_val):
                                # Continuation line
                                if txs and 'desc' in col_map and row[col_map['desc']]:
                                    desc_cont = str(row[col_map['desc']]).strip().replace("\n", " ")
                                    if desc_cont and not is_noise_line(desc_cont, "GENERIC") and not any(x in desc_cont.lower() for x in ("total", "carried forward", "brought forward", "page", "statement")):
                                        txs[-1]['Description'] += " " + desc_cont
                                continue
                                
                            desc_val = str(row[col_map['desc']] or "").strip().replace("\n", " ")
                            ref_val = str(row[col_map['ref']] or "").strip().replace("\n", " ") if 'ref' in col_map else ""
                            ref_match = re.search(r'\b\d{5,22}\b', ref_val)
                            if ref_match:
                                ref_val = ref_match.group(0)
                            elif not ref_val or ref_val == "-":
                                ref_val = ""
                                
                            # Resolve val_date_val supporting multi-line wrapped dates
                            val_date_val = ""
                            if 'val_date' in col_map and row[col_map['val_date']]:
                                raw_val_date_cell = str(row[col_map['val_date']] or "").strip()
                                val_date_lines = [l.strip() for l in raw_val_date_cell.split('\n') if l.strip()]
                                if val_date_lines:
                                    first_val_line = val_date_lines[0]
                                    date_patterns = [
                                        r'\b\d{2}[/\-.]\d{2}[/\-.]\d{2,4}\b',
                                        r'\b\d{2}[\s\-./][A-Za-z]{3}[\s\-./]\d{2,4}\b',
                                        r'\b\d{4}[\-./]\d{2}[\-./]\d{2}\b'
                                    ]
                                    if any(re.search(pat, first_val_line) for pat in date_patterns):
                                        val_date_val = first_val_line
                                    else:
                                        joined_val = " ".join(val_date_lines)
                                        val_date_val = re.sub(r'([\-./])\s+', r'\1', joined_val)
                            
                            debit_val = 0.0
                            credit_val = 0.0
                            
                            if 'debit' in col_map or 'credit' in col_map:
                                debit_val = clean_amount(row[col_map['debit']]) if 'debit' in col_map else 0.0
                                credit_val = clean_amount(row[col_map['credit']]) if 'credit' in col_map else 0.0
                            elif 'amount' in col_map:
                                amt = clean_amount(row[col_map['amount']])
                                type_str = str(row[col_map['type']]).lower() if 'type' in col_map else ""
                                if 'dr' in type_str or '-' in type_str or 'debit' in type_str:
                                    debit_val = amt
                                else:
                                    credit_val = amt
                                    
                            balance_val = clean_amount(row[col_map['balance']]) if 'balance' in col_map else 0.0
                            
                            # Handle merged columns in description cell (where amounts are grouped inside narration)
                            if debit_val == 0.0 and credit_val == 0.0:
                                desc_tokens = desc_val.strip().split()
                                likely_amounts = []
                                for tok in reversed(desc_tokens):
                                    tok_clean = tok.replace(",", "").strip()
                                    # Skip dates (e.g. 15.10.2024) or 6-digit cheque numbers (e.g. 289785)
                                    if re.match(r'^\d{1,2}[\-./]\d{2}[\-./]\d{2,4}$', tok) or re.match(r'^\d{6}$', tok_clean):
                                        break
                                    # Ensure it looks like an amount (e.g. has a decimal point)
                                    if '.' not in tok_clean:
                                        break
                                    try:
                                        val = float(tok_clean)
                                        likely_amounts.insert(0, (tok, val))
                                    except ValueError:
                                        break
                                
                                M = len(likely_amounts)
                                
                                # Diagnostic log
                                try:
                                    with open("debug_generic_parser.txt", "a", encoding="utf-8") as f_dbg:
                                        f_dbg.write(f"Date: {date_val} | M: {M} | Map: {str(col_map)} | Amounts: {str(likely_amounts)} | Desc: {desc_val}\n")
                                except Exception:
                                    pass
                                    
                                if M >= 2:
                                    # We found at least 2 amounts at the end
                                    amount1_str, amount1_val = likely_amounts[-2]
                                    amount2_str, amount2_val = likely_amounts[-1]
                                    
                                    # Check left-to-right order of debit and credit in col_map
                                    debit_idx = col_map.get('debit', -1)
                                    credit_idx = col_map.get('credit', -1)
                                    
                                    if debit_idx != -1 and credit_idx != -1 and debit_idx > credit_idx:
                                        credit_val = amount1_val
                                        debit_val = amount2_val
                                    else:
                                        # Default fallback (debit is first, credit is second)
                                        debit_val = amount1_val
                                        credit_val = amount2_val
                                        
                                    # Remove these amounts from description to clean it up
                                    desc_val = desc_val.replace(amount1_str, "").replace(amount2_str, "").strip()
                                    desc_val = re.sub(r'\s+$', '', desc_val)
                                elif M == 1:
                                    amount_str, amount_val = likely_amounts[0]
                                    if 'debit' in col_map and 'credit' not in col_map:
                                        debit_val = amount_val
                                    elif 'credit' in col_map and 'debit' not in col_map:
                                        credit_val = amount_val
                                    else:
                                        debit_val = amount_val
                                    desc_val = desc_val.replace(amount_str, "").strip()
                                    desc_val = re.sub(r'\s+$', '', desc_val)
                            
                            amount_val = credit_val if credit_val > 0 else debit_val
                            
                            txs.append({
                                'Date': date_val, 'Value Date': val_date_val, 'Description': desc_val,
                                'Ref / Cheque': ref_val, 'Amount': amount_val, 'Balance': balance_val,
                                'Debit': debit_val, 'Credit': credit_val
                            })
            if txs:
                return txs
    except Exception as e:
        print(f"Error in generic table extraction: {e}")
        
    # Fallback to line-by-line parsing if table extraction didn't work
    try:
        with open("debug_generic_parser.txt", "a", encoding="utf-8") as f_dbg:
            f_dbg.write("--- GENERIC FALLBACK LINE PARSER START ---\n")
    except Exception:
        pass

    raw_rows = []
    for line in lines:
        line_str = line.strip()
        if not line_str:
            continue
            
        # Match date formats like:
        # DD/MM/YYYY, DD-MM-YYYY, DD.MM.YYYY, DD/MM/YY
        # DD-Mmm-YYYY, DD Mmm YYYY, DD-Mmm-YY
        # YYYY-MM-DD
        # Can have an optional leading serial number
        date_pattern = r'^(?:\d+\s+)?(\d{2}[/\-.]\d{2}[/\-.]\d{2,4}|\d{2}[\s\-./][A-Za-z]{3}[\s\-./]\d{2,4}|\d{4}[\-./]\d{2}[\-./]\d{2})\s+'
        match = re.match(date_pattern, line_str)
        if match:
            date = match.group(1)
            tokens = line_str.split()
            
            try:
                with open("debug_generic_parser.txt", "a", encoding="utf-8") as f_dbg:
                    f_dbg.write(f"Match: {line_str} | Tokens: {str(tokens)}\n")
            except Exception:
                pass
                
            # Find the index of the date token
            date_token_idx = -1
            for idx, tok in enumerate(tokens):
                if date in tok or tok in date:
                    date_token_idx = idx
                    break
            desc_start_idx = date_token_idx + 1 if date_token_idx != -1 else 1
            
            if len(tokens) >= desc_start_idx + 1:
                has_indicator = tokens[-1].upper() in ('CR', 'DR')
                last_idx = -2 if has_indicator else -1
                
                amounts = []
                for idx in range(len(tokens) + last_idx, desc_start_idx - 1, -1):
                    tok = tokens[idx]
                    tok_clean = tok.replace(",", "").strip()
                    if re.match(r'^\d{1,2}[\-./]\d{2}[\-./]\d{2,4}$', tok) or re.match(r'^\d{6}$', tok_clean):
                        break
                    if '.' not in tok_clean:
                        break
                    try:
                        val = float(tok_clean)
                        amounts.insert(0, val)
                    except ValueError:
                        break
                
                debit_val = 0.0
                credit_val = 0.0
                balance_val = 0.0
                amount_val = 0.0
                
                num_amounts = len(amounts)
                if num_amounts >= 3:
                    debit_val = amounts[-3]
                    credit_val = amounts[-2]
                    balance_val = amounts[-1]
                    amount_val = credit_val if credit_val > 0.0 else debit_val
                    desc_tokens = tokens[desc_start_idx:-3 - (1 if has_indicator else 0)]
                elif num_amounts == 2:
                    amount_val = amounts[-2]
                    balance_val = amounts[-1]
                    desc_tokens = tokens[desc_start_idx:-2 - (1 if has_indicator else 0)]
                elif num_amounts == 1:
                    amount_val = amounts[0]
                    desc_tokens = tokens[desc_start_idx:-1 - (1 if has_indicator else 0)]
                else:
                    desc_tokens = tokens[desc_start_idx:]
                    
                desc_val = " ".join(desc_tokens)
                
                val_date_val = ""
                # Check if the first description token is a value date
                if len(desc_tokens) > 0 and re.match(r'^\d{2}[/\-.]\d{2}[/\-.]\d{2,4}$', desc_tokens[0]):
                    val_date_val = desc_tokens[0]
                    desc_val = " ".join(desc_tokens[1:])
                
                raw_rows.append((True, {
                    'Date': date, 'Value Date': val_date_val, 'Description': desc_val,
                    'Ref / Cheque': '', 'Amount': amount_val, 'Balance': balance_val,
                    'Debit': debit_val, 'Credit': credit_val
                }))
            else:
                raw_rows.append((False, line_str))
        else:
            raw_rows.append((False, line_str))

    # Second pass: split and merge prefix/suffix remarks
    merged_txs = []
    prefix_buffer = []
    in_prefix_zone = False
    prefix_keywords = ('inf/', 'neft', 'rtgs', 'bil/', 'imps', 'chq', 'chg', 'cash', 'by ', 'to ', 'upi', 'paytm', 'phonepe', 'gpay', 'amazonpay', 'ybl', 'axl', 'ibld', 'transfer', 'clg', 'int.', 'interest')

    for is_main, data in raw_rows:
        if is_main:
            tx = data
            if prefix_buffer:
                if tx['Description']:
                    tx['Description'] = " ".join(prefix_buffer) + " " + tx['Description']
                else:
                    tx['Description'] = " ".join(prefix_buffer)
                prefix_buffer = []
            merged_txs.append(tx)
            in_prefix_zone = False
        else:
            line_str = data
            if is_noise_line(line_str, "GENERIC"):
                continue
                
            is_prefix = line_str.lower().startswith(prefix_keywords)
            if is_prefix:
                in_prefix_zone = True
                
            if in_prefix_zone:
                prefix_buffer.append(line_str)
            else:
                if merged_txs:
                    if merged_txs[-1]['Description']:
                        merged_txs[-1]['Description'] += " " + line_str
                    else:
                        merged_txs[-1]['Description'] = line_str

    if not merged_txs:
        try:
            with open("debug_generic_parser.txt", "a", encoding="utf-8") as f_dbg:
                f_dbg.write("--- GENERIC BLOCK PARSER FALLBACK START ---\n")
        except Exception:
            pass
        print("[Parser] Standard parsing yielded no transactions. Running Heuristic Block Parser fallback...")
        merged_txs = parse_generic_blocks(lines)
        print(f"[Parser] Heuristic Block Parser: Extracted {len(merged_txs)} transactions successfully.")

    return merged_txs

# ---------------------------------------------------------------------------
# EasyOCR text extraction from scanned PDF pages
# ---------------------------------------------------------------------------
def extract_text_via_ocr(pdf_file_obj):
    """
    Convert each page of a scanned/image-only PDF to a greyscale image at
    300 DPI (via PyMuPDF) and run EasyOCR over it.  Word-boxes are grouped
    into text lines by their Y-coordinate proximity, preserving left-to-right
    reading order.  Returns (lines_list, full_text_str).
    """
    try:
        import fitz          # PyMuPDF
        import numpy as np
    except ImportError:
        raise ImportError("PyMuPDF is not installed. Run:  pip install pymupdf")

    reader = get_ocr_reader()

    if hasattr(pdf_file_obj, 'seek'):
        pdf_file_obj.seek(0)

    if hasattr(pdf_file_obj, 'read'):
        pdf_bytes = pdf_file_obj.read()
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    else:
        doc = fitz.open(pdf_file_obj)

    all_lines = []
    total_pages = len(doc)

    for page_num in range(total_pages):
        page = doc[page_num]
        print(f"[OCR] Processing page {page_num + 1}/{total_pages}...")

        # Render at 300 DPI in greyscale for best OCR accuracy
        mat = fitz.Matrix(300 / 72, 300 / 72)
        pix = page.get_pixmap(matrix=mat, colorspace=fitz.csGRAY)
        img_array = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w)

        # OCR — returns [(bbox, text, confidence), ...]
        results = reader.readtext(img_array, detail=1, paragraph=False)
        if not results:
            continue

        # Collect items above confidence threshold
        items = []
        for bbox, text, conf in results:
            if conf < 0.20:
                continue
            top_y  = bbox[0][1]   # top-left Y
            left_x = bbox[0][0]   # top-left X
            items.append((top_y, left_x, text))

        items.sort(key=lambda x: (x[0], x[1]))

        # Group into horizontal lines (items within 18 px share a row)
        page_lines = []
        current_items = []
        current_y = None
        Y_THRESHOLD = 18

        for top_y, left_x, text in items:
            if current_y is None or abs(top_y - current_y) <= Y_THRESHOLD:
                current_items.append((left_x, text))
                current_y = top_y if current_y is None else (current_y + top_y) / 2
            else:
                if current_items:
                    current_items.sort(key=lambda x: x[0])
                    page_lines.append('  '.join(t for _, t in current_items))
                current_items = [(left_x, text)]
                current_y = top_y

        if current_items:
            current_items.sort(key=lambda x: x[0])
            page_lines.append('  '.join(t for _, t in current_items))

        all_lines.extend(page_lines)

    doc.close()
    full_text = '\n'.join(all_lines)
    print(f"[OCR] Extracted {len(all_lines)} lines from scanned PDF.")
    return all_lines, full_text


# Parse a single PDF file
def process_pdf(pdf_path, stream_bytes=None, force_generic=False, enable_ocr=True):
    if stream_bytes:
        # Load PDF from memory stream
        pdf_file_obj = io.BytesIO(stream_bytes)
        filename = "uploaded_file.pdf"
    else:
        pdf_file_obj = pdf_path
        filename = os.path.basename(pdf_path)
        
    # Extract text from all pages
    full_text = ""
    lines = []
    has_chars = False
    with pdfplumber.open(pdf_file_obj) as pdf:
        for i, page in enumerate(pdf.pages):
            if page.chars:
                has_chars = True
            page_text = page.extract_text()
            if page_text:
                full_text += page_text + "\n"
                lines.extend(page_text.split("\n"))
                
    ocr_used = False

    if not lines:
        if not has_chars:
            if not enable_ocr:
                raise ValueError(f"File {filename} appears to be a scanned or image-based PDF. This software only supports text-based/digital PDFs.")
            # ── Scanned PDF detected — attempt EasyOCR ──────────────────────
            try:
                print(f"[OCR] No embedded text in '{filename}'. Switching to EasyOCR...")
                if hasattr(pdf_file_obj, 'seek'):
                    pdf_file_obj.seek(0)
                lines, full_text = extract_text_via_ocr(pdf_file_obj)
                ocr_used = True
            except ImportError as ie:
                raise ValueError(
                    f"'{filename}' is a scanned PDF. EasyOCR is not installed — "
                    f"please run: pip install easyocr pymupdf numpy\n({ie})"
                )
            except Exception as oe:
                raise ValueError(
                    f"'{filename}' is a scanned PDF and OCR failed: {oe}"
                )
            if not lines:
                raise ValueError(
                    f"'{filename}' is a scanned PDF but OCR could not extract any text. "
                    f"Please ensure the scan quality is sufficient."
                )
        else:
            return None

    # Detect Bank Type
    # OCR-processed files always use the generic parser (bank headers may be noisy)
    if ocr_used or force_generic:
        bank_type = "GENERIC"
    else:
        bank_type = detect_bank_type(full_text[:2000])
        
    try:
        with open("debug_uploaded_pdf.txt", "w", encoding="utf-8") as f_dbg:
            f_dbg.write(f"Filename: {filename}\n")
            f_dbg.write(f"Detected Bank Type: {bank_type}\n")
            f_dbg.write(f"Force Generic: {force_generic}\n")
            f_dbg.write("--- EXTRACTED TABLES START ---\n")
            if hasattr(pdf_file_obj, "seek"):
                pdf_file_obj.seek(0)
            with pdfplumber.open(pdf_file_obj) as pdf:
                for idx, page in enumerate(pdf.pages):
                    f_dbg.write(f"Page {idx+1} tables:\n")
                    tables = page.extract_tables()
                    for t_idx, table in enumerate(tables):
                        f_dbg.write(f"  Table {t_idx+1} (columns: {len(table[0]) if table else 0}):\n")
                        for r_idx, row in enumerate(table):
                            f_dbg.write(f"    Row {r_idx+1}: {str(row)}\n")
            f_dbg.write("--- EXTRACTED TABLES END ---\n")
            f_dbg.write("--- EXTRACTED TEXT START ---\n")
            f_dbg.write(full_text)
            f_dbg.write("--- EXTRACTED TEXT END ---\n")
    except Exception as e:
        pass
    
    parsers = {
        "KOTAK": parse_kotak,
        "IDFC": parse_idfc,
        "MODASA": parse_modasa,
        "SK": parse_sk,
        "MEHSANA": parse_mehsana,
        "UBI1": parse_ubi1,
        "UBI2": parse_ubi2
    }
    
    if bank_type not in parsers and bank_type not in ("ICICI", "BOB", "AXIS", "PRIME", "HDFC", "RBL", "SBI", "BOI", "UBI3", "GENERIC"):
        return None
        
    # Parse transaction rows
    if bank_type == "UBI3":
        if hasattr(pdf_file_obj, "seek"):
            pdf_file_obj.seek(0)
        raw_txs = parse_ubi3(pdf_file_obj, lines)
    elif bank_type == "ICICI":
        if hasattr(pdf_file_obj, "seek"):
            pdf_file_obj.seek(0)
        raw_txs = parse_icici(pdf_file_obj, lines)
    elif bank_type == "BOB":
        if hasattr(pdf_file_obj, "seek"):
            pdf_file_obj.seek(0)
        raw_txs = parse_bob(pdf_file_obj, lines)
    elif bank_type == "AXIS":
        if hasattr(pdf_file_obj, "seek"):
            pdf_file_obj.seek(0)
        raw_txs = parse_axis(pdf_file_obj)
    elif bank_type == "PRIME":
        if hasattr(pdf_file_obj, "seek"):
            pdf_file_obj.seek(0)
        raw_txs = parse_prime(pdf_file_obj)
    elif bank_type == "HDFC":
        if hasattr(pdf_file_obj, "seek"):
            pdf_file_obj.seek(0)
        raw_txs = parse_hdfc(pdf_file_obj, lines)
    elif bank_type == "RBL":
        if hasattr(pdf_file_obj, "seek"):
            pdf_file_obj.seek(0)
        raw_txs = parse_rbl(pdf_file_obj, lines)
    elif bank_type == "SBI":
        if hasattr(pdf_file_obj, "seek"):
            pdf_file_obj.seek(0)
        raw_txs = parse_sbi(pdf_file_obj, lines)
    elif bank_type == "BOI":
        if hasattr(pdf_file_obj, "seek"):
            pdf_file_obj.seek(0)
        raw_txs = parse_boi(pdf_file_obj, lines)
    elif bank_type == "GENERIC":
        if hasattr(pdf_file_obj, "seek"):
            pdf_file_obj.seek(0)
        raw_txs = parse_generic(pdf_file_obj, lines)
    else:
        raw_txs = parsers[bank_type](lines)
        
    # If bank-specific parsing yielded no transactions, fallback to the generic parser
    if not raw_txs:
        if hasattr(pdf_file_obj, "seek"):
            pdf_file_obj.seek(0)
        raw_txs = parse_generic(pdf_file_obj, lines)
    
    # Compute Debit/Credit columns
    is_reverse = (bank_type == "UBI1")
    txs = resolve_debits_credits(raw_txs, bank_type, is_reverse=is_reverse)
    
    # Standardize schema fields
    standard_txs = []
    for tx in txs:
        # Skip opening balance / brought forward rows with zero amount
        if tx['Amount'] == 0.0 and tx['Debit'] == 0.0 and tx['Credit'] == 0.0:
            continue
            
        ref_val = str(tx.get('Ref / Cheque', '') or '').strip()
        description = str(tx.get('Description', '') or '').strip()
        
        # If the cheque/ref number is empty or just a hyphen, try to extract it from description
        if not ref_val or ref_val == "-":
            # Extract 5 to 22 digit numbers
            ref_match = re.search(r'\b\d{5,22}\b', description)
            if ref_match:
                ref_val = ref_match.group(0)
            else:
                # Extract 4 digit numbers, excluding years 2000-2030
                for m in re.finditer(r'\b\d{4}\b', description):
                    val_str = m.group(0)
                    try:
                        year_val = int(val_str)
                        if 2000 <= year_val <= 2030:
                            continue
                    except ValueError:
                        pass
                    ref_val = val_str
                    break
        
        if ref_val and ref_val not in ("-", ""):
            if f"(Chq/Ref : {ref_val})" not in description:
                description = f"{description} (Chq/Ref : {ref_val})"
            
        standard_txs.append({
            'Bank': bank_type,
            'File Name': filename,
            'Data Source': 'EasyOCR' if ocr_used else 'Digital PDF',
            'Txn Date': parse_date(tx['Date']),
            'Value Date': parse_date(tx['Value Date']) if tx['Value Date'] else '',
            'Description': description,
            'Ref / Cheque No': ref_val,
            'Debit / Withdrawal': round(tx['Debit'], 2) if tx['Debit'] else 0.0,
            'Credit / Deposit': round(tx['Credit'], 2) if tx['Credit'] else 0.0,
            'Balance': round(tx['Balance'], 2)
        })
        
    return standard_txs

# API entry point for processing a list of files or file streams
def process_pdf_list(pdf_inputs, output_excel_path=None, force_generic=False):
    """
    pdf_inputs: List of tuples (filename, file_bytes_or_path)
    """
    all_data = []
    individual_sheets = {}
    
    for filename, file_data in pdf_inputs:
        try:
            if isinstance(file_data, bytes):
                txs = process_pdf(None, stream_bytes=file_data, force_generic=force_generic)
            else:
                txs = process_pdf(file_data, force_generic=force_generic)
        except ValueError as e:
            print(f"Error processing {filename}: {str(e)}")
            txs = None
            
        if txs:
            # Override filename in details
            for tx in txs:
                tx['File Name'] = filename
            df = pd.DataFrame(txs)
            all_data.extend(txs)
            
            sheet_name = os.path.splitext(filename)[0][:30]
            # Replace special characters for excel sheet naming constraints
            sheet_name = re.sub(r'[\\/\?\*\[\]]', '_', sheet_name)
            individual_sheets[sheet_name] = df
            
    if not all_data:
        return None
        
    consolidated_df = pd.DataFrame(all_data)
    
    column_order = [
        'Bank',
        'File Name',
        'Txn Date',
        'Value Date',
        'Description',
        'Ref / Cheque No',
        'Debit / Withdrawal',
        'Credit / Deposit',
        'Balance'
    ]
    
    # Ensure all columns exist
    for col in column_order:
        if col not in consolidated_df.columns:
            consolidated_df[col] = ''
            
    consolidated_df = consolidated_df[column_order]
    consolidated_df = consolidated_df.sort_values(by=['Bank', 'Txn Date'], ascending=[True, True])
    
    # Helper to convert YYYY-MM-DD to DD-MM-YYYY
    def to_dd_mm_yyyy(val):
        if not val:
            return ""
        try:
            dt = datetime.strptime(val, "%Y-%m-%d")
            return dt.strftime("%d-%m-%Y")
        except ValueError:
            return val

    consolidated_df['Txn Date'] = consolidated_df['Txn Date'].apply(to_dd_mm_yyyy)
    consolidated_df['Value Date'] = consolidated_df['Value Date'].apply(to_dd_mm_yyyy)
    
    if output_excel_path:
        with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
            consolidated_df.to_excel(writer, sheet_name="Consolidated", index=False)
            format_worksheet(writer.sheets["Consolidated"])
            for sheet_name, df_sheet in individual_sheets.items():
                for col in column_order:
                    if col not in df_sheet.columns:
                        df_sheet[col] = ''
                df_sheet_ordered = df_sheet[column_order].copy()
                df_sheet_ordered['Txn Date'] = df_sheet_ordered['Txn Date'].apply(to_dd_mm_yyyy)
                df_sheet_ordered['Value Date'] = df_sheet_ordered['Value Date'].apply(to_dd_mm_yyyy)
                df_sheet_ordered.to_excel(writer, sheet_name=sheet_name, index=False)
                format_worksheet(writer.sheets[sheet_name])
            
            # Generate and write Tally Voucher sheet
            df_tally = generate_tally_voucher_df(consolidated_df)
            df_tally.to_excel(writer, sheet_name="Tally Voucher", index=False)
            format_tally_worksheet(writer.sheets["Tally Voucher"])
        return output_excel_path
    else:
        output_buffer = io.BytesIO()
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            consolidated_df.to_excel(writer, sheet_name="Consolidated", index=False)
            format_worksheet(writer.sheets["Consolidated"])
            for sheet_name, df_sheet in individual_sheets.items():
                for col in column_order:
                    if col not in df_sheet.columns:
                        df_sheet[col] = ''
                df_sheet_ordered = df_sheet[column_order].copy()
                df_sheet_ordered['Txn Date'] = df_sheet_ordered['Txn Date'].apply(to_dd_mm_yyyy)
                df_sheet_ordered['Value Date'] = df_sheet_ordered['Value Date'].apply(to_dd_mm_yyyy)
                df_sheet_ordered.to_excel(writer, sheet_name=sheet_name, index=False)
                format_worksheet(writer.sheets[sheet_name])
            
            # Generate and write Tally Voucher sheet
            df_tally = generate_tally_voucher_df(consolidated_df)
            df_tally.to_excel(writer, sheet_name="Tally Voucher", index=False)
            format_tally_worksheet(writer.sheets["Tally Voucher"])
        output_buffer.seek(0)
        return output_buffer.getvalue()

def parse_excel_file(file_bytes_or_path, filename, selected_sheet=None):
    """
    Parses transactions from uploaded Excel/xlsm sheets (including TALLY EXCEL).
    Auto-detects the transaction sheet and maps columns dynamically.
    """
    import io
    import uuid
    import pandas as pd
    
    if isinstance(file_bytes_or_path, bytes):
        file_obj = io.BytesIO(file_bytes_or_path)
    else:
        file_obj = file_bytes_or_path
        
    xl = pd.ExcelFile(file_obj)
    sheet_names = xl.sheet_names
    
    # Resolve target sheet: use selected sheet if valid, otherwise fallback to heuristics
    target_sheet = None
    if selected_sheet and selected_sheet in sheet_names:
        target_sheet = selected_sheet
    else:
        for name in ('Run Code', 'DRCR', 'WD', 'Transactions'):
            if name in sheet_names:
                target_sheet = name
                break
        if not target_sheet:
            target_sheet = sheet_names[0]
        
    # Read the sheet
    # For 'Run Code', the headers are at row 5 (0-indexed row 4)
    header_idx = 0
    if target_sheet == 'Run Code':
        # Inspect first 10 rows to locate the header row
        df_temp = xl.parse(target_sheet, header=None, nrows=10)
        for idx, row in df_temp.iterrows():
            row_vals = [str(x).lower() for x in row if pd.notna(x)]
            if any('voucher date' in x or 'voucher number' in x for x in row_vals):
                header_idx = idx
                break
                
    df = xl.parse(target_sheet, header=header_idx)
    
    # Standardize column mapping
    col_map = {}
    for idx, col in enumerate(df.columns):
        col_clean = str(col).lower().replace("\n", " ").replace("_", " ").replace(".", "").strip()
        
        # Voucher/Date
        if 'voucher date' in col_clean or col_clean == 'date':
            col_map['date'] = col
        # Narration
        elif 'voucher narration' in col_clean or 'narration' in col_clean or 'description' in col_clean or 'particular' in col_clean:
            col_map['desc'] = col
        # Dr/Cr (evaluated early to avoid collision with ledger amount)
        elif 'dr/cr' in col_clean or 'drcr' in col_clean or 'dr cr' in col_clean or 'dr/ cr' in col_clean:
            col_map['drcr'] = col
        # Debit/Withdrawal
        elif 'withdrawal' in col_clean or 'debit' in col_clean:
            col_map['debit'] = col
        # Credit/Deposit
        elif 'deposit' in col_clean or 'credit' in col_clean:
            col_map['credit'] = col
        # Ledger Amount / Amount
        elif 'ledger amount' in col_clean or 'amount' in col_clean:
            col_map['amount'] = col
        # Balance
        elif 'balance' in col_clean or col_clean == 'bal':
            col_map['balance'] = col
        # Ref/Cheque
        elif 'ref' in col_clean or 'cheque' in col_clean or 'chq' in col_clean:
            col_map['ref'] = col
        # Bank Name
        elif 'bank' in col_clean:
            col_map['bank'] = col
            
    # Process rows
    raw_txs = []
    for idx, row in df.iterrows():
        # Date value check
        date_col = col_map.get('date')
        if not date_col or pd.isna(row[date_col]):
            continue
            
        date_str = str(row[date_col]).strip()
        if not re.search(r'\d', date_str):
            continue
            
        # Parse timestamp string safely if pandas parsed it as datetime
        if isinstance(row[date_col], pd.Timestamp):
            date_str = row[date_col].strftime("%d-%m-%Y")
        else:
            # Clean date string
            date_str = date_str.split(" ")[0].strip()
            
        desc_val = str(row[col_map['desc']]) if col_map.get('desc') and pd.notna(row[col_map['desc']]) else ""
        ref_val = str(row[col_map['ref']]) if col_map.get('ref') and pd.notna(row[col_map['ref']]) else ""
        
        debit_val = 0.0
        if col_map.get('debit') and pd.notna(row[col_map['debit']]):
            try:
                debit_val = float(str(row[col_map['debit']]).replace(",", ""))
            except ValueError:
                pass
                
        credit_val = 0.0
        if col_map.get('credit') and pd.notna(row[col_map['credit']]):
            try:
                credit_val = float(str(row[col_map['credit']]).replace(",", ""))
            except ValueError:
                pass
                
        amount_val = 0.0
        if col_map.get('amount') and pd.notna(row[col_map['amount']]):
            try:
                amount_val = float(str(row[col_map['amount']]).replace(",", ""))
            except ValueError:
                pass
                
        balance_val = 0.0
        if col_map.get('balance') and pd.notna(row[col_map['balance']]):
            try:
                balance_val = float(str(row[col_map['balance']]).replace(",", ""))
            except ValueError:
                pass
                
        drcr_val = str(row[col_map['drcr']]).strip().upper() if col_map.get('drcr') and pd.notna(row[col_map['drcr']]) else ""
        
        # If amount_val is provided but not debit/credit, resolve it
        if amount_val > 0.0 and debit_val == 0.0 and credit_val == 0.0:
            if 'DR' in drcr_val:
                debit_val = amount_val
            elif 'CR' in drcr_val:
                credit_val = amount_val
            else:
                # Default assume debit
                debit_val = amount_val
                
        bank_name = str(row[col_map['bank']]).strip() if col_map.get('bank') and pd.notna(row[col_map['bank']]) else "Excel Import"
        
        # Generate an id for client-side matching
        txn_id = str(uuid.uuid4())
        
        raw_txs.append({
            'id': txn_id,
            'Bank': bank_name,
            'File Name': filename,
            'Data Source': 'Excel Import',
            'Txn Date': date_str,
            'Value Date': date_str,
            'Description': desc_val,
            'Ref / Cheque No': ref_val,
            'Debit / Withdrawal': debit_val,
            'Credit / Deposit': credit_val,
            'Balance': balance_val,
            'Amount': amount_val if amount_val > 0.0 else (debit_val if debit_val > 0.0 else credit_val)
        })
        
    return raw_txs

# Command line compilation script execution
def main():
    pdf_dir = "Bank Format pdf"
    pdf_pattern = os.path.join(pdf_dir, "*.pdf")
    pdf_files = glob.glob(pdf_pattern)
    
    if not pdf_files:
        print(f"No PDF files found in '{pdf_dir}' folder!")
        sys.exit(1)
        
    print(f"Found {len(pdf_files)} PDF statements to convert.\n")
    
    pdf_inputs = [(os.path.basename(f), f) for f in pdf_files]
    output_excel = "bank_transactions.xlsx"
    
    result = process_pdf_list(pdf_inputs, output_excel_path=output_excel)
    if result:
        print("\nConversion successfully completed!")
        print(f"Output saved in the workspace root directory: {os.path.abspath(output_excel)}")
    else:
        print("No transactions were parsed successfully.")
        sys.exit(1)

if __name__ == "__main__":
    main()
