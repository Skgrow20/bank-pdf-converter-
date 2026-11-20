import io
import os
import re
import html
from datetime import datetime
from flask import Flask, request, render_template, send_file, jsonify, session, redirect, url_for, flash, abort, g
import pandas as pd
from bank_pdf_to_excel import process_pdf, generate_tally_voucher_df, format_tally_worksheet, format_worksheet, format_summary_worksheet, parse_excel_file
from openpyxl.styles import Font, PatternFill, Alignment
import user_manager
import device_manager
import uuid

import sys

if getattr(sys, 'frozen', False):
    # Running in a PyInstaller bundle
    template_folder = os.path.join(sys._MEIPASS, 'templates')
    app = Flask(__name__, template_folder=template_folder)
else:
    app = Flask(__name__)

app.secret_key = 'bank_statement_converter_auth_secret_key_19.06.26'

# Server-side in-memory cache for preview transactions to prevent Flask cookie size overflow
_PREVIEW_CACHE = {}

def get_preview_cache_id():
    cid = session.get('preview_cache_id')
    if not cid:
        cid = uuid.uuid4().hex
        session['preview_cache_id'] = cid
    return cid

@app.errorhandler(500)
def internal_error(e):
    import traceback
    tb = traceback.format_exc()
    err_text = tb if tb and tb.strip() != "NoneType: None" else str(e)
    return f"""
    <!DOCTYPE html>
    <html>
    <head><title>SK Software - Diagnostic</title></head>
    <body style="font-family: monospace; background: #0f172a; color: #f8fafc; padding: 30px;">
        <h2 style="color: #ef4444;">Server Error Diagnostic (500)</h2>
        <p style="color: #94a3b8;">An error occurred while handling this request:</p>
        <pre style="background: #1e293b; padding: 20px; border-radius: 8px; color: #fca5a5; overflow-x: auto; font-size: 13px;">{err_text}</pre>
        <br>
        <a href="/" style="display: inline-block; padding: 10px 20px; background: #4f46e5; color: white; text-decoration: none; border-radius: 6px;">Return to Home</a>
    </body>
    </html>
    """, 500




@app.before_request
def check_auth_and_device():
    # Check if request is local (127.0.0.1, ::1, or localhost)
    is_local = (request.remote_addr in ['127.0.0.1', '::1', 'localhost'])
    
    token = None
    if is_local:
        # Retrieve physical BIOS UUID
        bios_uuid = device_manager.get_bios_uuid()
        if bios_uuid:
            token = "bios_" + bios_uuid
            
    # Fallback to cookie-based token if remote or bios check failed
    if not token:
        token = request.cookies.get('device_token')
        if not token:
            # Generate a temporary token to set in after_request
            token = 'device_' + uuid.uuid4().hex
            g.new_device_token = token
    
    g.device_token = token
    
    # Auto-authorize the first connected device so the admin is never locked out
    if not device_manager.has_any_authorized_devices() and request.endpoint != 'static':
        device_manager.authorize_device(token, name="Primary PC (Auto-Authorized)")
    
    # 2. Check if device is authorized
    is_auth = device_manager.is_device_authorized(token)
    
    # Register in pending queue if unauthorized
    if not is_auth:
        device_manager.add_pending_request(token, request.remote_addr)
    
    # Allow public access to Google verification, robots.txt, and static assets
    if request.path.startswith('/google') or request.path in ['/robots.txt', '/favicon.ico']:
        return None

    # Allow access to login, logout, unauthorized error page, register, verify_otp, and static assets
    if request.endpoint in ['login', 'logout', 'device_unauthorized', 'static', 'register', 'verify_otp', 'robots_txt', 'google_verification_file'] or not request.endpoint:
        return None
        
    if not is_auth:
        return redirect(url_for('device_unauthorized'))
        
    # 3. Check login session
    if 'username' not in session:
        return redirect(url_for('login'))

@app.after_request
def set_device_cookie(response):
    new_token = getattr(g, 'new_device_token', None)
    if new_token:
        # Set persistent cookie for 5 years
        response.set_cookie('device_token', new_token, max_age=5*365*24*60*60, httponly=True)
    return response

@app.context_processor
def inject_user():
    return {
        'current_user': {
            'username': session.get('username'),
            'role': session.get('role')
        }
    }



@app.route('/')
def index():
    return render_template('dashboard.html')


@app.route('/pdf_to_tally')
def pdf_to_tally():
    return render_template('index.html')


@app.route('/excel_to_tally', methods=['GET', 'POST'])
def excel_to_tally():
    if request.method == 'POST':
        if 'files' not in request.files:
            return render_template('excel_index.html', error='No file part in the request')
            
        uploaded_files = request.files.getlist('files')
        if not uploaded_files or uploaded_files[0].filename == '':
            return render_template('excel_index.html', error='No files selected')
            
        # Get selected sheet name if passed
        selected_sheet = request.form.get('selected_sheet')
        if selected_sheet == '':
            selected_sheet = None
            
        all_transactions = []
        for file in uploaded_files:
            if file.filename.lower().endswith(('.xlsx', '.xlsm')):
                try:
                    file_bytes = file.read()
                    txs = parse_excel_file(file_bytes, file.filename, selected_sheet=selected_sheet)
                    all_transactions.extend(txs)
                except Exception as e:
                    return render_template('excel_index.html', error=f'Failed to parse Excel file {file.filename}: {str(e)}')
                    
        if not all_transactions:
            return render_template('excel_index.html', error='No transactions found in the uploaded Excel worksheets. Make sure the sheets are named Run Code, DRCR, WD, or Transactions.')
            
        # Sort transactions
        all_transactions.sort(key=lambda x: (x['Bank'], x['Txn Date'] or ""))
        
        return render_template('preview.html', transactions=all_transactions)
        
    return render_template('excel_index.html')


@app.route('/get_excel_sheets', methods=['POST'])
def get_excel_sheets():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
        
    try:
        import openpyxl
        # Read Excel workbook read-only to fetch sheet names instantly
        wb = openpyxl.load_workbook(file, read_only=True, keep_vba=False)
        sheet_names = wb.sheetnames
        wb.close()
        return jsonify({'sheets': sheet_names})
    except Exception as e:
        return jsonify({'error': str(e)}), 500



@app.route('/preview', methods=['GET', 'POST'])
def preview():
    if request.method == 'GET':
        cid = session.get('preview_cache_id')
        cached_txs = _PREVIEW_CACHE.get(cid) if cid else None
        if cached_txs:
            return render_template('preview.html', transactions=cached_txs)
        return redirect(url_for('pdf_to_tally'))

    if 'files' not in request.files:
        return render_template('index.html', error='No file part in the request')
        
    uploaded_files = request.files.getlist('files')
    if not uploaded_files or uploaded_files[0].filename == '':
        return render_template('index.html', error='No files selected')
        
    format_mode = request.form.get('format_mode', 'auto')
    enable_ocr = (request.form.get('enable_ocr') == 'true')
    
    pdf_inputs = []
    for file in uploaded_files:
        if file.filename.lower().endswith('.pdf'):
            file_bytes = file.read()
            pdf_inputs.append((file.filename, file_bytes))
            
    if not pdf_inputs:
        return render_template('index.html', error='No valid PDF files were uploaded')
        
    try:
        all_transactions = []
        scanned_files = []
        for filename, file_data in pdf_inputs:
            try:
                txs = process_pdf(None, stream_bytes=file_data, force_generic=(format_mode == 'generic'), enable_ocr=enable_ocr)
                
                if txs:
                    for tx in txs:
                        tx['File Name'] = filename
                    all_transactions.extend(txs)
            except ValueError as e:
                if "scanned" in str(e).lower():
                    scanned_files.append(filename)
                else:
                    raise e
                    
        if scanned_files:
            error_msg = f"Failed to parse. The following file(s) appear to be scanned or image-based PDFs: {', '.join(scanned_files)}. This software only supports text-based/digital PDFs. Please upload statements downloaded directly from NetBanking."
            return render_template('index.html', error=error_msg)
            
        if not all_transactions:
            return render_template('index.html', error='Failed to parse any transactions from the uploaded statements. Please check if they are supported formats.')
            
        # Sort chronologically by YYYY-MM-DD first
        all_transactions.sort(key=lambda x: (x['Bank'], x['Txn Date'] or ""))
        
        # Convert date fields from YYYY-MM-DD to DD-MM-YYYY for display in browser preview
        for tx in all_transactions:
            if tx['Txn Date']:
                try:
                    dt = datetime.strptime(tx['Txn Date'], "%Y-%m-%d")
                    tx['Txn Date'] = dt.strftime("%d-%m-%Y")
                except ValueError:
                    pass
            if tx['Value Date']:
                try:
                    dt = datetime.strptime(tx['Value Date'], "%Y-%m-%d")
                    tx['Value Date'] = dt.strftime("%d-%m-%Y")
                except ValueError:
                    pass
            
        cid = get_preview_cache_id()
        _PREVIEW_CACHE[cid] = all_transactions
        session.pop('cached_transactions', None)
        return render_template('preview.html', transactions=all_transactions)
    except Exception as e:
        return render_template('index.html', error=f'An unexpected error occurred during processing: {str(e)}')

def populate_tally_xlsm(template_path, transactions, ledger_mapping=None):
    """
    Populates sheets DRCR, WD, and Run Code of a macro-enabled Excel workbook template
    while preserving VBA code, shapes, and layouts.
    """
    import openpyxl
    if ledger_mapping is None:
        ledger_mapping = {}
        
    wb = openpyxl.load_workbook(template_path, keep_vba=True)
    
    # 1. Populate DRCR sheet
    if 'DRCR' in wb.sheetnames:
        ws = wb['DRCR']
        # Clear existing data from row 2 onwards
        if ws.max_row >= 2:
            for r in range(2, ws.max_row + 1):
                for c in range(1, 5):
                    ws.cell(r, c).value = None
                    
        for idx, tx in enumerate(transactions):
            row = idx + 2
            ws.cell(row, 1).value = tx.get('Date') or tx.get('Txn Date', '')
            ws.cell(row, 2).value = tx.get('Narration') or tx.get('Description', '')
            
            # Amount
            debit_str = str(tx.get('Debit (Rs)') or tx.get('Debit / Withdrawal') or 0.0).replace(",", "")
            credit_str = str(tx.get('Credit (Rs)') or tx.get('Credit / Deposit') or 0.0).replace(",", "")
            debit = 0.0
            credit = 0.0
            try:
                if debit_str and debit_str != '-':
                    debit = float(debit_str)
            except ValueError:
                pass
            try:
                if credit_str and credit_str != '-':
                    credit = float(credit_str)
            except ValueError:
                pass
                
            amount = debit if debit > 0.0 else credit
            ws.cell(row, 3).value = amount
            ws.cell(row, 4).value = 'Dr' if debit > 0.0 else 'Cr'
            
    # 2. Populate WD sheet
    if 'WD' in wb.sheetnames:
        ws = wb['WD']
        # Clear existing data from row 2 onwards
        if ws.max_row >= 2:
            for r in range(2, ws.max_row + 1):
                for c in range(1, 5):
                    ws.cell(r, c).value = None
                    
        for idx, tx in enumerate(transactions):
            row = idx + 2
            ws.cell(row, 1).value = tx.get('Date') or tx.get('Txn Date', '')
            ws.cell(row, 2).value = tx.get('Narration') or tx.get('Description', '')
            
            debit_str = str(tx.get('Debit (Rs)') or tx.get('Debit / Withdrawal') or 0.0).replace(",", "")
            credit_str = str(tx.get('Credit (Rs)') or tx.get('Credit / Deposit') or 0.0).replace(",", "")
            debit = 0.0
            credit = 0.0
            try:
                if debit_str and debit_str != '-':
                    debit = float(debit_str)
            except ValueError:
                pass
            try:
                if credit_str and credit_str != '-':
                    credit = float(credit_str)
            except ValueError:
                pass
                
            ws.cell(row, 3).value = debit if debit > 0.0 else None
            ws.cell(row, 4).value = credit if credit > 0.0 else None
            
    # Calculate financial summary metrics
    total_debit = 0.0
    total_credit = 0.0
    for tx in transactions:
        deb_s = str(tx.get('Debit (Rs)') or tx.get('Debit / Withdrawal') or 0.0).replace(",", "").replace("-", "")
        cred_s = str(tx.get('Credit (Rs)') or tx.get('Credit / Deposit') or 0.0).replace(",", "").replace("-", "")
        try:
            total_debit += float(deb_s) if deb_s else 0.0
        except ValueError:
            pass
        try:
            total_credit += float(cred_s) if cred_s else 0.0
        except ValueError:
            pass

    opening_balance = 0.0
    closing_balance = 0.0
    if transactions:
        first_tx = transactions[0]
        bal_s = str(first_tx.get('Balance (Rs)') or first_tx.get('Balance') or 0.0).replace(",", "")
        d1_s = str(first_tx.get('Debit (Rs)') or first_tx.get('Debit / Withdrawal') or 0.0).replace(",", "")
        c1_s = str(first_tx.get('Credit (Rs)') or first_tx.get('Credit / Deposit') or 0.0).replace(",", "")
        try:
            first_bal = float(bal_s) if bal_s else 0.0
        except ValueError:
            first_bal = 0.0
        try:
            first_deb = float(d1_s) if d1_s else 0.0
        except ValueError:
            first_deb = 0.0
        try:
            first_cred = float(c1_s) if c1_s else 0.0
        except ValueError:
            first_cred = 0.0

        opening_balance = first_bal + first_deb - first_cred
        closing_balance = opening_balance - total_debit + total_credit

    # 3. Populate Run Code sheet
    if 'Run Code' in wb.sheetnames:
        ws = wb['Run Code']
        
        # Populate Financial Summary KPI header cells
        ws['F1'].value = round(opening_balance, 2)
        ws['F2'].value = round(total_debit, 2)
        ws['F3'].value = round(total_credit, 2)
        ws['F4'].value = round(closing_balance, 2)

        # Clear existing data from row 6 onwards
        if ws.max_row >= 6:
            for r in range(6, ws.max_row + 1):
                for c in range(1, 8):
                    ws.cell(r, c).value = None
                    
        row_idx = 6
        for idx, tx in enumerate(transactions):
            date = tx.get('Date') or tx.get('Txn Date', '')
            narration = tx.get('Narration') or tx.get('Description', '')
            bank = tx.get('Bank') or 'Bank Account'
            mapped_ledger = ledger_mapping.get(bank) or 'Suspense'
            
            debit_str = str(tx.get('Debit (Rs)') or tx.get('Debit / Withdrawal') or 0.0).replace(",", "")
            credit_str = str(tx.get('Credit (Rs)') or tx.get('Credit / Deposit') or 0.0).replace(",", "")
            debit = 0.0
            credit = 0.0
            try:
                if debit_str and debit_str != '-':
                    debit = float(debit_str)
            except ValueError:
                pass
            try:
                if credit_str and credit_str != '-':
                    credit = float(credit_str)
            except ValueError:
                pass
                
            amount = debit if debit > 0.0 else credit
            vtype = 'Payment' if debit > 0.0 else 'Receipt'
            
            # Row 1: Suspense side
            ws.cell(row_idx, 1).value = date
            ws.cell(row_idx, 2).value = vtype
            ws.cell(row_idx, 3).value = idx + 1
            ws.cell(row_idx, 4).value = mapped_ledger
            ws.cell(row_idx, 5).value = amount
            ws.cell(row_idx, 6).value = narration
            ws.cell(row_idx, 7).value = 'Dr' if debit > 0.0 else 'Cr'
            row_idx += 1
            
            # Row 2: Bank side
            ws.cell(row_idx, 1).value = date
            ws.cell(row_idx, 2).value = vtype
            ws.cell(row_idx, 3).value = idx + 1
            ws.cell(row_idx, 4).value = bank
            ws.cell(row_idx, 5).value = amount
            ws.cell(row_idx, 6).value = narration
            ws.cell(row_idx, 7).value = 'Cr' if debit > 0.0 else 'Dr'
            row_idx += 1
            
    return wb

# API Route: Convert JSON preview array back to multi-sheet Excel workbook
@app.route('/download', methods=['POST'])
def download():
    data = request.json
    if not data or 'transactions' not in data:
        return jsonify({'error': 'No data provided'}), 400
        
    transactions = data['transactions']
    if not transactions:
        return jsonify({'error': 'No transactions to compile'}), 400
        
    try:
        # Check if macro-enabled export specifically requested
        export_format = data.get('format', 'excel')
        template_path = os.path.join("EXCEL", "TALLY EXCEL.xlsm")
        if export_format == 'macro' and os.path.exists(template_path):
            print(f"[Excel] Macro export requested. Populating {template_path}...")
            ledger_mapping = data.get('ledger_mapping', {})
            wb = populate_tally_xlsm(template_path, transactions, ledger_mapping)
            
            output_buffer = io.BytesIO()
            wb.save(output_buffer)
            output_buffer.seek(0)
            
            response = send_file(
                output_buffer,
                mimetype='application/vnd.ms-excel.sheet.macroEnabled.12',
                as_attachment=True,
                download_name='TALLY EXCEL.xlsm'
            )
            response.headers["Access-Control-Expose-Headers"] = "Content-Disposition"
            return response
            
        df = pd.DataFrame(transactions)
        
        old_columns = ['Bank', 'File Name', 'Txn Date', 'Value Date', 'Description', 'Ref / Cheque No', 'Debit / Withdrawal', 'Credit / Deposit', 'Balance']
        for col in old_columns:
            if col not in df.columns:
                df[col] = ''
                
        # Sort Consolidated transactions correctly by parsing DD-MM-YYYY first, fallback to YYYY-MM-DD
        temp_dates = pd.to_datetime(df['Txn Date'], format='%d-%m-%Y', errors='coerce')
        nat_mask = temp_dates.isna() & df['Txn Date'].astype(str).str.strip().ne('')
        if nat_mask.any():
            temp_dates[nat_mask] = pd.to_datetime(df.loc[nat_mask, 'Txn Date'], format='%Y-%m-%d', errors='coerce')
            
        df['temp_date'] = temp_dates
        df_sorted = df.sort_values(by=['Bank', 'temp_date'], ascending=[True, True])
        df_sorted = df_sorted.drop(columns=['temp_date'])
        
        # Ensure all date columns in the Excel sheets are formatted as DD-MM-YYYY
        def to_dd_mm_yyyy(val):
            if not val:
                return ""
            val_str = str(val).strip()
            try:
                dt = datetime.strptime(val_str, "%Y-%m-%d")
                return dt.strftime("%d-%m-%Y")
            except ValueError:
                pass
            try:
                datetime.strptime(val_str, "%d-%m-%Y")
                return val_str
            except ValueError:
                pass
            return val_str

        df_sorted['Txn Date'] = df_sorted['Txn Date'].apply(to_dd_mm_yyyy)
        df_sorted['Value Date'] = df_sorted['Value Date'].apply(to_dd_mm_yyyy)
        
        df['Txn Date'] = df['Txn Date'].apply(to_dd_mm_yyyy)
        df['Value Date'] = df['Value Date'].apply(to_dd_mm_yyyy)
        
        # Generate Tally Voucher DataFrame (using the old schema columns before renaming)
        df_tally = generate_tally_voucher_df(df_sorted)
        
        # 1. Compute global consecutive running balances across all chronologically sorted transactions
        if not df_sorted.empty:
            first_tx = df_sorted.iloc[0]
            first_credit = float(str(first_tx.get('Credit / Deposit') or 0.0).replace(',', '')) if first_tx.get('Credit / Deposit') else 0.0
            first_debit = float(str(first_tx.get('Debit / Withdrawal') or 0.0).replace(',', '')) if first_tx.get('Debit / Withdrawal') else 0.0
            first_bal = float(str(first_tx.get('Balance') or 0.0).replace(',', '')) if first_tx.get('Balance') else 0.0
            global_opening_bal = first_bal - first_credit + first_debit
            
            running_bal = global_opening_bal
            calculated_balances = []
            for _, row in df_sorted.iterrows():
                deb = float(str(row.get('Debit / Withdrawal') or 0.0).replace(',', '')) if row.get('Debit / Withdrawal') else 0.0
                cred = float(str(row.get('Credit / Deposit') or 0.0).replace(',', '')) if row.get('Credit / Deposit') else 0.0
                running_bal = running_bal - deb + cred
                calculated_balances.append(round(running_bal, 2))
                
            df_sorted['Balance'] = calculated_balances
            df['Balance'] = calculated_balances
            global_closing_bal = calculated_balances[-1] if calculated_balances else global_opening_bal
        else:
            global_opening_bal = 0.0
            global_closing_bal = 0.0

        # 2. Generate Summary DataFrame
        summary_rows = []
        file_groups = df_sorted.groupby('File Name', sort=False)
        for file_name, group in file_groups:
            total_credits = group['Credit / Deposit'].apply(lambda x: float(str(x).replace(',', '')) if x else 0.0).sum()
            total_debits = group['Debit / Withdrawal'].apply(lambda x: float(str(x).replace(',', '')) if x else 0.0).sum()
            bank_name = group.iloc[0].get('Bank', '')
            
            f_first = group.iloc[0]
            f_last = group.iloc[-1]
            f_credit_1 = float(str(f_first.get('Credit / Deposit') or 0.0).replace(',', '')) if f_first.get('Credit / Deposit') else 0.0
            f_debit_1 = float(str(f_first.get('Debit / Withdrawal') or 0.0).replace(',', '')) if f_first.get('Debit / Withdrawal') else 0.0
            f_bal_1 = float(str(f_first.get('Balance') or 0.0).replace(',', '')) if f_first.get('Balance') else 0.0
            f_open = f_bal_1 - f_credit_1 + f_debit_1
            f_close = float(str(f_last.get('Balance') or 0.0).replace(',', ''))
            
            summary_rows.append({
                'Source File': file_name,
                'Bank': bank_name,
                'Opening Balance': round(f_open, 2),
                'Total Credits': round(total_credits, 2),
                'Total Debits': round(total_debits, 2),
                'Closing Balance': round(f_close, 2)
            })
            
        total_credits = df_sorted['Credit / Deposit'].apply(lambda x: float(str(x).replace(',', '')) if x else 0.0).sum()
        total_debits = df_sorted['Debit / Withdrawal'].apply(lambda x: float(str(x).replace(',', '')) if x else 0.0).sum()
        
        if len(summary_rows) == 1:
            summary_rows[0]['Opening Balance'] = round(global_opening_bal, 2)
            summary_rows[0]['Closing Balance'] = round(global_closing_bal, 2)
            
        summary_rows.append({
            'Source File': 'Total',
            'Bank': 'All',
            'Opening Balance': round(global_opening_bal, 2),
            'Total Credits': round(total_credits, 2),
            'Total Debits': round(total_debits, 2),
            'Closing Balance': round(global_closing_bal, 2)
        })
            
        df_summary = pd.DataFrame(summary_rows)
        
        # Rename columns to match new schema
        rename_map = {
            'Txn Date': 'Date',
            'Description': 'Narration',
            'Debit / Withdrawal': 'Debit (Rs)',
            'Credit / Deposit': 'Credit (Rs)',
            'Balance': 'Balance (Rs)',
            'Ref / Cheque No': 'Chq / Ref No',
            'File Name': 'Source File'
        }
        df_sorted = df_sorted.rename(columns=rename_map)
        df = df.rename(columns=rename_map)
        
        # Preserve per-transaction Data Source (EasyOCR / Digital PDF)
        # Fall back to 'Digital PDF' only if the column is missing entirely
        if 'Data Source' not in df_sorted.columns:
            df_sorted['Data Source'] = 'Digital PDF'
        else:
            df_sorted['Data Source'] = df_sorted['Data Source'].fillna('Digital PDF')
        if 'Data Source' not in df.columns:
            df['Data Source'] = 'Digital PDF'
        else:
            df['Data Source'] = df['Data Source'].fillna('Digital PDF')
        
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
        
        for col in column_order:
            if col not in df_sorted.columns:
                df_sorted[col] = ''
            if col not in df.columns:
                df[col] = ''
                
        df_transactions = df_sorted[column_order].copy()
        df_reordered = df[column_order].copy()
        
        output_buffer = io.BytesIO()
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            # 1. Write Transactions sheet
            df_transactions.to_excel(writer, sheet_name="Transactions", index=False)
            format_worksheet(writer.sheets["Transactions"])
            
            # 2. Write Summary sheet
            df_summary.to_excel(writer, sheet_name="Summary", index=False)
            format_summary_worksheet(writer.sheets["Summary"])
            
            # 3. Write individual file sheet(s) grouped under "Per File" naming convention
            unique_files = df_reordered['Source File'].unique()
            if len(unique_files) == 1:
                df_file = df_reordered[df_reordered['Source File'] == unique_files[0]]
                df_file.to_excel(writer, sheet_name="Per File", index=False)
                format_worksheet(writer.sheets["Per File"])
            else:
                for filename in unique_files:
                    df_file = df_reordered[df_reordered['Source File'] == filename]
                    sheet_name = f"Per File - {os.path.splitext(filename)[0]}"[:30]
                    sheet_name = re.sub(r'[\\/\?\*\[\]]', '_', sheet_name)
                    df_file.to_excel(writer, sheet_name=sheet_name, index=False)
                    format_worksheet(writer.sheets[sheet_name])
                
            # 4. Write Tally Voucher sheet
            df_tally.to_excel(writer, sheet_name="Tally Voucher", index=False)
            format_tally_worksheet(writer.sheets["Tally Voucher"])
                
        output_buffer.seek(0)
        return send_file(
            output_buffer,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='bank_transactions_compiled.xlsx'
        )
    except Exception as e:
        return jsonify({'error': f'Failed to generate Excel: {str(e)}'}), 500


def to_tally_date(date_str):
    if not date_str:
        return ""
    # Standardize separators to dashes
    clean_date = str(date_str).replace(".", "-").replace("/", "-").strip()
    try:
        dt = datetime.strptime(clean_date, "%d-%m-%Y")
        return dt.strftime("%Y%m%d")
    except ValueError:
        pass
    try:
        dt = datetime.strptime(clean_date, "%Y-%m-%d")
        return dt.strftime("%Y%m%d")
    except ValueError:
        pass
    return clean_date.replace("-", "")


def get_tally_url():
    """Retrieve configured Tally URL from session, file, environment, or default."""
    url = session.get('tally_url')
    if not url:
        config_path = os.path.join(os.path.dirname(__file__), 'tally_config.json')
        if os.path.exists(config_path):
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    cfg = json.load(f)
                    url = cfg.get('tally_url')
            except Exception:
                pass
    if not url:
        url = os.environ.get('TALLY_URL', 'http://localhost:9000')
    if not url.startswith('http://') and not url.startswith('https://'):
        url = 'http://' + url
    return url.rstrip('/')

@app.route('/tally_config', methods=['GET', 'POST'])
def tally_config():
    config_path = os.path.join(os.path.dirname(__file__), 'tally_config.json')
    if request.method == 'POST':
        data = request.json or {}
        host = data.get('host', '').strip()
        port = str(data.get('port', '')).strip()
        if host:
            host_clean = host.strip()
            if host_clean.startswith(('http://', 'https://')):
                tally_url = host_clean.rstrip('/')
            elif any(d in host_clean.lower() for d in ('ngrok', 'trycloudflare', 'localtunnel', '.app', '.io', '.me')) and ':' not in host_clean:
                tally_url = f"https://{host_clean.rstrip('/')}"
            else:
                port_val = port if port else '9000'
                clean_no_proto = re.sub(r'^https?://', '', host_clean).rstrip('/')
                if ':' in clean_no_proto:
                    tally_url = f"http://{clean_no_proto}"
                else:
                    tally_url = f"http://{clean_no_proto}:{port_val}"
        else:
            tally_url = data.get('tally_url', '').strip() or 'http://localhost:9000'
            
        if not tally_url.startswith(('http://', 'https://')):
            tally_url = 'http://' + tally_url

            
        session['tally_url'] = tally_url
        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump({'tally_url': tally_url}, f, indent=2)
        except Exception as e:
            print(f"[Tally Config] Warning: Could not save config file: {e}")
            
        return tally_status()
        
    current_url = get_tally_url()
    return jsonify({'tally_url': current_url})

def send_tally_xml_request(tally_url, xml_data, timeout=8):
    """Send XML request to Tally Prime using built-in urllib (no external requests dependency)."""
    import urllib.request
    import urllib.error
    req = urllib.request.Request(
        tally_url,
        data=xml_data.encode('utf-8') if isinstance(xml_data, str) else xml_data,
        headers={
            'Content-Type': 'text/xml;charset=utf-8',
            'ngrok-skip-browser-warning': 'true',
            'User-Agent': 'Mozilla/5.0'
        },
        method='POST'
    )
    with urllib.request.urlopen(req, timeout=timeout) as response:
        return response.getcode(), response.read().decode('utf-8', errors='ignore')

@app.route('/tally_status', methods=['GET'])
def tally_status():
    tally_url = get_tally_url()
    
    # Request to get current active company
    active_company_xml = """<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>GetActiveCompany</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION NAME="GetActiveCompany" ISINITIALIZE="Yes">
            <TYPE>Company</TYPE>
            <NATIVEMETHOD>Name</NATIVEMETHOD>
            <FILTER>IsCurrentCompany</FILTER>
          </COLLECTION>
          <SYSTEM TYPE="Formulae" NAME="IsCurrentCompany">$Name = $$CurrentCompany</SYSTEM>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>"""

    try:
        code, res_text = send_tally_xml_request(tally_url, active_company_xml, timeout=5)
        if code == 200:
            # Parse the company name from the response XML
            company_match = re.search(r'<COMPANY[^>]*>.*?<NAME[^>]*>(.*?)</NAME>.*?</COMPANY>', res_text, re.DOTALL)
            if not company_match:
                company_match = re.search(r'<NAME[^>]*>(.*?)</NAME>', res_text)
                
            if company_match:
                company_name = html.unescape(company_match.group(1).strip())
                return jsonify({'connected': True, 'company': company_name, 'tally_url': tally_url})
            else:
                return jsonify({'connected': True, 'company': None, 'tally_url': tally_url})
        else:
            return jsonify({'connected': False, 'tally_url': tally_url, 'error': f'Tally at {tally_url} returned status code {code}'})
    except Exception as e:
        err_msg = str(e)
        if 'Connection refused' in err_msg or 'WinError 10061' in err_msg:
            err_msg = f'Could not connect to Tally Prime at {tally_url}. Please ensure Tally Prime is running with Port 9000 enabled.'
        return jsonify({'connected': False, 'tally_url': tally_url, 'error': err_msg})


@app.route('/tally_bank_ledgers', methods=['GET'])
def tally_bank_ledgers():
    tally_url = get_tally_url()
    
    # Request to get bank ledgers
    bank_ledgers_xml = """<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>GetBankLedgers</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION NAME="GetBankLedgers" ISINITIALIZE="Yes">
            <TYPE>Ledger</TYPE>
            <NATIVEMETHOD>Name</NATIVEMETHOD>
            <FILTER>BankGroupFilter</FILTER>
          </COLLECTION>
          <SYSTEM TYPE="Formulae" NAME="BankGroupFilter">$Parent = "Bank Accounts" or $Parent = "Bank OCC A/c" or $Parent = "Bank OD A/c"</SYSTEM>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>"""

    try:
        code, res_text = send_tally_xml_request(tally_url, bank_ledgers_xml, timeout=6)
        if code == 200:
            # Parse names using regex
            ledger_names = re.findall(r'<NAME[^>]*>(.*?)</NAME>', res_text)
            
            # Clean and filter out duplicates
            cleaned_names = []
            for name in ledger_names:
                name_clean = html.unescape(name.strip())
                if name_clean and name_clean not in cleaned_names:
                    cleaned_names.append(name_clean)
            return jsonify({'success': True, 'ledgers': sorted(cleaned_names)})
        else:
            return jsonify({'success': False, 'error': f'Tally returned status {code}'})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})



@app.route('/import_to_tally', methods=['POST'])
def import_to_tally():
    if session.get('role') == 'viewer':
        return jsonify({'error': 'Viewer accounts do not have permission to import to Tally.'}), 403
        
    data = request.json
    if not data or 'transactions' not in data:
        return jsonify({'error': 'No data provided'}), 400
        
    transactions = data['transactions']
    if not transactions:
        return jsonify({'error': 'No transactions to import'}), 400
        
    ledger_mapping = data.get('ledger_mapping', {})
    tally_url = get_tally_url()
    
    # Extract unique mapped bank ledger names
    unique_tally_ledgers = set()
    for tx in transactions:
        bank = tx.get('Bank') or tx.get('Bank Name') or 'Bank Account'
        tally_ledger = ledger_mapping.get(bank) or bank
        unique_tally_ledgers.add(tally_ledger)
        
    # Generate ledger master creation XML messages dynamically
    bank_messages = []
    for bank_name in sorted(list(unique_tally_ledgers)):
        bank_messages.append(f"""                <TALLYMESSAGE xmlns:UDF="TallyUDF">
                    <LEDGER NAME="{bank_name}" ACTION="Create">
                        <NAME>{bank_name}</NAME>
                        <PARENT>Bank Accounts</PARENT>
                    </LEDGER>
                </TALLYMESSAGE>""")
    bank_messages_str = "\n".join(bank_messages)
    
    # Create Ledgers XML
    ledgers_xml = f"""<ENVELOPE>
    <HEADER>
        <TALLYREQUEST>Import Data</TALLYREQUEST>
    </HEADER>
    <BODY>
        <IMPORTDATA>
            <REQUESTDESC>
                <REPORTNAME>All Masters</REPORTNAME>
            </REQUESTDESC>
            <REQUESTDATA>
                <TALLYMESSAGE xmlns:UDF="TallyUDF">
                    <LEDGER NAME="Suspense" ACTION="Create">
                        <NAME>Suspense</NAME>
                        <PARENT>Suspense Accounts</PARENT>
                    </LEDGER>
                </TALLYMESSAGE>
{bank_messages_str}
            </REQUESTDATA>
        </IMPORTDATA>
    </BODY>
</ENVELOPE>"""

    headers = {
        'Content-Type': 'text/xml;charset=utf-8',
        'ngrok-skip-browser-warning': 'true'
    }
    try:
        send_tally_xml_request(tally_url, ledgers_xml, timeout=10)
    except Exception as e:
        return jsonify({'error': f'Failed to send masters to Tally: {str(e)}'}), 500


    # Build Vouchers XML
    voucher_entries = []
    for i, tx in enumerate(transactions, 1):
        date_str = tx.get('Txn Date') or tx.get('Date') or ''
        tally_date = to_tally_date(date_str)
        narration = tx.get('Description') or tx.get('Narration') or ''
        bank_name = tx.get('Bank') or tx.get('Bank Name') or 'Bank Account'
        tally_ledger_name = ledger_mapping.get(bank_name) or bank_name
        
        debit = 0.0
        if tx.get('Debit / Withdrawal') is not None and str(tx.get('Debit / Withdrawal')).strip():
            try:
                debit = float(str(tx.get('Debit / Withdrawal')).replace(',', ''))
            except ValueError:
                pass
        if tx.get('Debit (Rs)') is not None and str(tx.get('Debit (Rs)')).strip():
            try:
                debit = float(str(tx.get('Debit (Rs)')).replace(',', ''))
            except ValueError:
                pass
                
        credit = 0.0
        if tx.get('Credit / Deposit') is not None and str(tx.get('Credit / Deposit')).strip():
            try:
                credit = float(str(tx.get('Credit / Deposit')).replace(',', ''))
            except ValueError:
                pass
        if tx.get('Credit (Rs)') is not None and str(tx.get('Credit (Rs)')).strip():
            try:
                credit = float(str(tx.get('Credit (Rs)')).replace(',', ''))
            except ValueError:
                pass
                
        if debit > 0:
            vtype = "Payment"
            amount = debit
            entry = f"""                <TALLYMESSAGE xmlns:UDF="TallyUDF">
                    <VOUCHER ACTION="Create">
                        <DATE>{tally_date}</DATE>
                        <VOUCHERTYPENAME>{vtype}</VOUCHERTYPENAME>
                        <VOUCHERNUMBER>{i}</VOUCHERNUMBER>
                        <NARRATION>{narration}</NARRATION>
                        <ALLLEDGERENTRIES.LIST>
                            <LEDGERNAME>Suspense</LEDGERNAME>
                            <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
                            <AMOUNT>-{amount:.2f}</AMOUNT>
                        </ALLLEDGERENTRIES.LIST>
                        <ALLLEDGERENTRIES.LIST>
                            <LEDGERNAME>{tally_ledger_name}</LEDGERNAME>
                            <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
                            <AMOUNT>{amount:.2f}</AMOUNT>
                        </ALLLEDGERENTRIES.LIST>
                    </VOUCHER>
                </TALLYMESSAGE>"""
        else:
            vtype = "Receipt"
            amount = credit
            entry = f"""                <TALLYMESSAGE xmlns:UDF="TallyUDF">
                    <VOUCHER ACTION="Create">
                        <DATE>{tally_date}</DATE>
                        <VOUCHERTYPENAME>{vtype}</VOUCHERTYPENAME>
                        <VOUCHERNUMBER>{i}</VOUCHERNUMBER>
                        <NARRATION>{narration}</NARRATION>
                        <ALLLEDGERENTRIES.LIST>
                            <LEDGERNAME>Suspense</LEDGERNAME>
                            <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
                            <AMOUNT>{amount:.2f}</AMOUNT>
                        </ALLLEDGERENTRIES.LIST>
                        <ALLLEDGERENTRIES.LIST>
                            <LEDGERNAME>{tally_ledger_name}</LEDGERNAME>
                            <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
                            <AMOUNT>-{amount:.2f}</AMOUNT>
                        </ALLLEDGERENTRIES.LIST>
                    </VOUCHER>
                </TALLYMESSAGE>"""
        voucher_entries.append(entry)
        
    vouchers_xml_body = "\n".join(voucher_entries)
    vouchers_xml = f"""<ENVELOPE>
    <HEADER>
        <TALLYREQUEST>Import Data</TALLYREQUEST>
    </HEADER>
    <BODY>
        <IMPORTDATA>
            <REQUESTDESC>
                <REPORTNAME>Vouchers</REPORTNAME>
            </REQUESTDESC>
            <REQUESTDATA>
{vouchers_xml_body}
            </REQUESTDATA>
        </IMPORTDATA>
    </BODY>
</ENVELOPE>"""

    try:
        code, res_text = send_tally_xml_request(tally_url, vouchers_xml, timeout=90)
        created_match = re.search(r'<CREATED>(\d+)</CREATED>', res_text)
        errors_match = re.search(r'<ERRORS>(\d+)</ERRORS>', res_text)
        
        created = int(created_match.group(1)) if created_match else None
        errors = int(errors_match.group(1)) if errors_match else None
        
        if created is None or errors is None:
            err_line_match = re.search(r'<LINEERROR>(.*?)</LINEERROR>', res_text)
            if err_line_match:
                return jsonify({'error': f'Tally Error: {err_line_match.group(1)}'}), 422
            clean_res = re.sub(r'<[^>]+>', ' ', res_text).strip()
            return jsonify({'error': f'Tally Response: {clean_res or "Please open a Company in Tally Prime first."}'}), 422
                
        if errors and errors > 0:
            return jsonify({
                'error': f'Imported with errors. Vouchers Created: {created}, Errors: {errors}. Please check the Tally.imp log file in your Tally directory.'
            }), 422
            
        return jsonify({
            'message': f'Successfully imported {created} vouchers directly into Tally Prime!'
        })
    except Exception as e:
        return jsonify({'error': f'Failed to import vouchers to Tally: {str(e)}'}), 500


@app.route('/login', methods=['GET', 'POST'])
def login():
    if 'username' in session:
        return redirect(url_for('index'))
        
    error = None
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')
        
        user = user_manager.authenticate_user(username, password)
        if user:
            is_admin = (user['role'] == 'admin')
            device_token = request.cookies.get('device_token') or getattr(g, 'device_token', None)
            
            # If not admin, verify if this device is authorized
            if not is_admin and not device_manager.is_device_authorized(device_token):
                error = "This PC/device is not authorized. Please ask your administrator to authorize it."
            else:
                # If admin, auto-authorize the device on login
                if is_admin and device_token and not device_manager.is_device_authorized(device_token):
                    device_manager.authorize_device(device_token, name=f"Authorized PC ({username})")
                    
                session['username'] = user['username']
                session['role'] = user['role']
                return redirect(url_for('index'))
        else:
            error = "Invalid username or password."
            
    return render_template('login.html', error=error)


@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))


@app.route('/device_unauthorized')
def device_unauthorized():
    token = request.cookies.get('device_token') or getattr(g, 'device_token', '')
    device_id = device_manager.get_device_id(token)
    return render_template('device_unauthorized.html', device_id=device_id)


@app.route('/devices/approve', methods=['POST'])
def approve_device():
    if session.get('role') != 'admin':
        abort(403)
        
    device_id = request.form.get('device_id', '').strip()
    name = request.form.get('name', 'Remote PC').strip()
    
    if not device_id:
        flash("Device ID cannot be empty.", "error")
        return redirect(url_for('users'))
        
    success = device_manager.authorize_by_device_id(device_id, name)
    if success:
        flash(f"Device '{name}' ({device_id}) authorized successfully.", "success")
    else:
        # Fallback to authorize direct token if they pasted a raw token
        success = device_manager.authorize_device(device_id, name)
        if success:
            flash(f"Device '{name}' authorized successfully.", "success")
        else:
            flash(f"Failed to find or authorize device '{device_id}'.", "error")
            
    return redirect(url_for('users'))


@app.route('/devices/revoke', methods=['POST'])
def revoke_device_route():
    if session.get('role') != 'admin':
        abort(403)
        
    token = request.form.get('token', '').strip()
    if not token:
        flash("Token cannot be empty.", "error")
        return redirect(url_for('users'))
        
    success = device_manager.revoke_device(token)
    if success:
        flash("Device authorization revoked successfully.", "success")
    else:
        flash("Failed to revoke device access.", "error")
        
    return redirect(url_for('users'))


@app.route('/devices/clear_pending', methods=['POST'])
def clear_pending_devices_route():
    if session.get('role') != 'admin':
        abort(403)
        
    device_manager.clear_all_pending()
    flash("All pending device authorization requests cleared successfully.", "success")
    return redirect(url_for('users'))


@app.route('/devices/reset_all', methods=['POST'])
def reset_all_devices_route():
    if session.get('role') != 'admin':
        abort(403)
        
    current_token = getattr(g, 'device_token', None)
    device_manager.reset_all_devices(keep_token=current_token)
    flash("All pending requests cleared and unauthorized devices reset. Your current PC remains authorized.", "success")
    return redirect(url_for('users'))


@app.route('/users')
def users():
    if session.get('role') != 'admin':
        abort(403)
    all_users = user_manager.list_users()
    authorized_devices = device_manager.list_authorized_devices()
    pending_requests = device_manager.list_pending_requests()
    email_config = user_manager.load_email_config()
    gsheet_config = user_manager.load_gsheet_config()
    return render_template('users.html', 
                           users=all_users, 
                           authorized_devices=authorized_devices, 
                           pending_requests=pending_requests,
                           email_config=email_config,
                           gsheet_config=gsheet_config)


@app.route('/users/add', methods=['POST'])
def add_user_route():
    if session.get('role') != 'admin':
        abort(403)
        
    username = request.form.get('username', '').strip()
    password = request.form.get('password', '')
    role = request.form.get('role', 'viewer')
    full_name = request.form.get('full_name', '').strip()
    mobile = request.form.get('mobile', '').strip()
    email = request.form.get('email', '').strip()
    
    device_token = getattr(g, 'device_token', None)
    device_id = device_manager.get_device_id(device_token) or "Admin Console"
    ip_addr = request.remote_addr or "127.0.0.1"
    
    success, msg = user_manager.add_user(
        username=username, 
        password=password, 
        role=role, 
        full_name=full_name, 
        mobile=mobile, 
        email=email, 
        device_id=device_id, 
        ip_address=ip_addr
    )
    if success:
        flash(f"{msg}", 'success')
    else:
        flash(msg, 'error')
        
    return redirect(url_for('users'))


@app.route('/admin/change_credentials', methods=['POST'])
def change_admin_credentials():
    if session.get('role') != 'admin':
        abort(403)
        
    current_username = session.get('username')
    new_username = request.form.get('new_username', '').strip()
    current_password = request.form.get('current_password', '')
    new_password = request.form.get('new_password', '')
    confirm_password = request.form.get('confirm_password', '')
    
    if new_password != confirm_password:
        flash("New password and confirm password do not match.", "error")
        return redirect(url_for('users'))
        
    ok, msg = user_manager.change_admin_credentials(
        current_user=current_username,
        new_username=new_username,
        current_password=current_password,
        new_password=new_password
    )
    if ok:
        session['username'] = new_username
        flash(f"✅ {msg} Your new login username is '{new_username}'.", "success")
    else:
        flash(f"❌ {msg}", "error")
        
    return redirect(url_for('users'))


@app.route('/users/download_excel')
def download_users_excel():
    if session.get('role') != 'admin':
        abort(403)
    flash("User accounts are stored in Google Sheets Cloud instead of local Excel.", "info")
    return redirect(url_for('users'))


@app.route('/users/delete', methods=['POST'])
def delete_user_route():
    if session.get('role') != 'admin':
        abort(403)
        
    username = request.form.get('username', '').strip()
    success, msg = user_manager.delete_user(username, session.get('username'))
    if success:
        flash(f"{msg}", 'success')
    else:
        flash(msg, 'error')
        
    return redirect(url_for('users'))


@app.route('/admin/email_config', methods=['POST'])
def update_email_config():
    if session.get('role') != 'admin':
        abort(403)
        
    server = request.form.get('smtp_server', '').strip()
    port_str = request.form.get('smtp_port', '').strip()
    sender = request.form.get('sender_email', '').strip()
    password = request.form.get('sender_password', '').strip()
    admin_email = request.form.get('admin_email', '').strip()
    
    try:
        port = int(port_str)
    except ValueError:
        port = 587
        
    config = {
        "smtp_server": server,
        "smtp_port": port,
        "sender_email": sender,
        "sender_password": password,
        "admin_email": admin_email
    }
    
    if user_manager.save_email_config(config):
        flash("SMTP email configuration updated successfully.", "success")
    else:
        flash("Failed to save email configuration.", "error")
        
    return redirect(url_for('users'))


@app.route('/admin/google_sheets_config', methods=['POST'])
def update_google_sheets_config():
    if session.get('role') != 'admin':
        abort(403)
        
    action = request.form.get('action', 'save')
    webhook_url = request.form.get('webhook_url', '').strip()
    enabled = ('enabled' in request.form) or (action in ['test', 'sync_all'] and bool(webhook_url))
    
    # If a webhook URL was submitted, save and auto-enable it
    if webhook_url:
        user_manager.save_gsheet_config({
            "webhook_url": webhook_url,
            "enabled": True
        })
        
    if action == 'test':
        ok, msg = user_manager.test_google_sheet_connection(override_url=webhook_url)
        if ok:
            flash(f"✅ Google Sheets Test Success: {msg}", "success")
        else:
            flash(f"❌ Google Sheets Test Failed: {msg}", "error")
        return redirect(url_for('users'))
        
    if action == 'sync_all':
        ok, msg = user_manager.sync_all_users_to_google_sheet()
        if ok:
            flash(f"✅ {msg}", "success")
        else:
            flash(f"❌ {msg}", "error")
        return redirect(url_for('users'))
        
    config = {
        "webhook_url": webhook_url,
        "enabled": enabled
    }
    
    if user_manager.save_gsheet_config(config):
        flash("Google Sheets cloud configuration updated successfully.", "success")
    else:
        flash("Failed to save Google Sheets configuration.", "error")
        
    return redirect(url_for('users'))


@app.route('/register', methods=['GET', 'POST'])
def register():
    if 'username' in session:
        return redirect(url_for('index'))
        
    error = None
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')
        full_name = request.form.get('full_name', '').strip()
        mobile = request.form.get('mobile', '').strip()
        email = request.form.get('email', '').strip()
        
        if not username:
            error = "Username cannot be empty."
        elif not password or len(password) < 4:
            error = "Password must be at least 4 characters long."
        else:
            # Capture device token and authorize immediately
            device_token = getattr(g, 'device_token', None)
            device_id = device_manager.get_device_id(device_token) or "Web Browser"
            ip_addr = request.remote_addr or "127.0.0.1"
            
            if device_token:
                device_manager.authorize_device(device_token, name=f"{full_name or username} (PC)")
                
            success, msg = user_manager.add_user(
                username=username, 
                password=password, 
                role='operator', 
                full_name=full_name, 
                mobile=mobile, 
                email=email, 
                device_id=device_id, 
                ip_address=ip_addr
            )
            if success:
                flash("Account created successfully! You can now log in.", "success")
                return redirect(url_for('login'))
            else:
                error = msg
                
    return render_template('register.html', error=error)


@app.route('/verify_otp', methods=['GET', 'POST'])
def verify_otp():
    if 'username' in session:
        return redirect(url_for('index'))
        
    username = request.args.get('username', '').strip()
    if not username:
        return redirect(url_for('register'))
        
    error = None
    if request.method == 'POST':
        input_otp = request.form.get('otp', '').strip()
        success, msg = user_manager.verify_otp_and_activate(username, input_otp)
        if success:
            flash("Registration completed! Details saved to Google Sheets cloud.", 'success')
            return redirect(url_for('login'))
        else:
            error = msg
            
    return render_template('verify_otp.html', username=username, error=error)


@app.route('/robots.txt')
def robots_txt():
    content = "User-agent: *\nAllow: /\nDisallow: /admin\nDisallow: /users\n"
    return content, 200, {'Content-Type': 'text/plain'}


@app.route('/google9c968ba59091574f.html')
def google_verification_file():
    return "google-site-verification: google9c968ba59091574f.html", 200, {'Content-Type': 'text/html'}


if __name__ == '__main__':
    import webbrowser
    from threading import Timer
    
    def open_browser():
        webbrowser.open("http://127.0.0.1:5000")
        
    # Only open the browser on the primary thread (prevent double-launch from Werkzeug reloader)
    if not os.environ.get("WERKZEUG_RUN_MAIN"):
        Timer(1.0, open_browser).start()
        
    # Disable debug mode and reloader when running as a compiled standalone executable
    is_frozen = getattr(sys, 'frozen', False)
    port = int(os.environ.get("PORT", 5000))
    host = "0.0.0.0" if os.environ.get("PORT") else "127.0.0.1"
    app.run(host=host, port=port, debug=not is_frozen)
