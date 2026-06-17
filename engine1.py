import pandas as pd
import numpy as np
import io
import logging

logger = logging.getLogger(__name__)

def parse_tally_file(file_path):
    # Read first 20 rows to find header
    raw = pd.read_excel(file_path, nrows=20, header=None)
    header_row = None
    
    for i in range(raw.shape[0]):
        row_str = ' '.join([str(x).lower() for x in raw.iloc[i] if pd.notna(x)])
        if any(keyword in row_str for keyword in ['voucher no', 'bill no', 'invoice number', 'invoice no']):
            header_row = i
            break
            
    if header_row is None:
        raise ValueError("Could not find Tally header row (looking for 'Voucher No', 'Bill No', or 'Invoice Number')")
        
    file_path.seek(0)
    df = pd.read_excel(file_path, header=header_row)
    
    # Store raw columns to check for warnings
    raw_columns = [str(c).strip() for c in df.columns]
    
    col_map = {}
    for col in df.columns:
        col_lower = str(col).lower().strip()
        if 'invoice number' in col_lower or 'invoice no' in col_lower or 'voucher no' in col_lower or 'bill no' in col_lower:
            if 'invoice_no' not in col_map.values():
                col_map[col] = 'invoice_no'
        elif col_lower == 'date' or 'invoice date' in col_lower:
            col_map[col] = 'invoice_date'
        elif 'gstin' in col_lower:
            col_map[col] = 'gstin'
        elif 'party' in col_lower:
            col_map[col] = 'party_name_tally'
        elif 'gross minus discount' in col_lower or 'taxable value' in col_lower:
            col_map[col] = 'taxable_value'
        elif 'gst amount' in col_lower or 'tax amount' in col_lower or 'total tax' in col_lower:
            col_map[col] = 'tax_amount'
        elif 'integrated tax' in col_lower or col_lower == 'igst':
            col_map[col] = 'igst'
        elif 'central tax' in col_lower or col_lower == 'cgst':
            col_map[col] = 'cgst'
        elif 'state/ut tax' in col_lower or 'state tax' in col_lower or col_lower == 'sgst':
            col_map[col] = 'sgst'
            
    df.rename(columns=col_map, inplace=True)
    
    if 'tax_amount' not in df.columns:
        # Try to sum igst, cgst, sgst if available
        tax_sum = pd.Series(0.0, index=df.index)
        found_taxes = False
        for t_col in ['igst', 'cgst', 'sgst']:
            if t_col in df.columns:
                found_taxes = True
                temp = df[t_col].astype(str).str.replace(',', '', regex=False).str.replace('₹', '', regex=False).str.strip()
                tax_sum += pd.to_numeric(temp, errors='coerce').fillna(0)
        if found_taxes:
            df['tax_amount'] = tax_sum
            
    if 'invoice_no' not in df.columns:
        df['invoice_no'] = ''
        
    df['invoice_no'] = df['invoice_no'].astype(str).str.strip().replace(['nan', 'NA'], '')
    
    # 1. Identify primary rows to generate NO_INV for blank invoices
    is_primary = pd.Series(False, index=df.index)
    if 'invoice_date' in df.columns:
        is_primary |= df['invoice_date'].astype(str).str.strip().replace(['nan', 'NA', 'NaT', 'None'], '') != ''
    if 'party_name_tally' in df.columns:
        is_primary |= df['party_name_tally'].astype(str).str.strip().replace(['nan', 'NA', 'None'], '') != ''
        
    blank_primary_mask = is_primary & (df['invoice_no'] == '')
    df.loc[blank_primary_mask, 'invoice_no'] = ['NO_INV_' + str(i) for i in df.index[blank_primary_mask]]
    
    # 2. Forward fill all header details so secondary rows inherit them safely
    df['invoice_no'] = df['invoice_no'].replace('', np.nan).ffill()
    
    if 'gstin' in df.columns:
        df['gstin'] = df['gstin'].astype(str).str.strip().str.upper()
        df['gstin'] = df['gstin'].replace(['NAN', 'NA', 'NONE', ''], np.nan).ffill().fillna('UNREGISTERED')
        
    if 'party_name_tally' in df.columns:
        df['party_name_tally'] = df['party_name_tally'].astype(str).str.strip().replace(['nan', 'NA', 'None', ''], np.nan).ffill()
        
    if 'invoice_date' in df.columns:
        df['invoice_date'] = df['invoice_date'].astype(str).str.strip().replace(['nan', 'NA', 'NaT', 'None', ''], np.nan).ffill()
        
    # 3. Extract Round Off to add to parent, then drop Round Off
    product_cols = [c for c in df.columns if 'product' in str(c).lower() or 'other account' in str(c).lower()]
    round_off_mask = pd.Series(False, index=df.index)
    for col in product_cols:
        round_off_mask |= df[col].astype(str).str.lower().str.contains('round off', na=False)
        
    special_rows = df[round_off_mask].copy()
    
    agg_cols = []
    for col in ['taxable_value', 'tax_amount']:
        if col in special_rows.columns:
            special_rows[col] = special_rows[col].astype(str).str.replace(',', '', regex=False).str.replace('₹', '', regex=False).str.strip()
            special_rows[col] = pd.to_numeric(special_rows[col], errors='coerce').fillna(0)
            agg_cols.append(col)
            
    if agg_cols and not special_rows.empty:
        special_agg = special_rows.groupby('invoice_no')[agg_cols].sum()
    else:
        special_agg = pd.DataFrame()
        
    df = df[~round_off_mask]
    
    if not special_agg.empty:
        adjustments = special_agg.reset_index()
        if 'gstin' in df.columns:
            adjustments['gstin'] = adjustments['invoice_no'].map(df.groupby('invoice_no')['gstin'].first())
        if 'party_name_tally' in df.columns:
            adjustments['party_name_tally'] = adjustments['invoice_no'].map(df.groupby('invoice_no')['party_name_tally'].first())
        if 'invoice_date' in df.columns:
            adjustments['invoice_date'] = adjustments['invoice_no'].map(df.groupby('invoice_no')['invoice_date'].first())
            
        df = pd.concat([df, adjustments], ignore_index=True)
        
    df = df[df['invoice_no'].notna()]
    df = df[~df['invoice_no'].astype(str).str.contains('Total', na=False, case=False)]
    
    return df, raw_columns

def aggregate_tally(df):
    if 'invoice_no' not in df.columns:
        df['invoice_no'] = ''
    if 'gstin' not in df.columns:
        df['gstin'] = ''
        
    df['invoice_no'] = df['invoice_no'].astype(str).str.strip()
    df['gstin'] = df['gstin'].astype(str).str.strip().str.upper()
    
    for col in ['taxable_value', 'tax_amount']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(',', '', regex=False).str.replace('₹', '', regex=False).str.strip()
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        else:
            df[col] = 0.0
            
    if 'invoice_date' in df.columns:
        df['invoice_date'] = pd.to_datetime(df['invoice_date'], errors='coerce', dayfirst=True)
        
    # CRITICAL: Group by invoice_no and gstin
    agg_funcs = {
        'taxable_value': 'sum',
        'tax_amount': 'sum'
    }
    if 'invoice_date' in df.columns:
        agg_funcs['invoice_date'] = 'first'
    if 'party_name_tally' in df.columns:
        agg_funcs['party_name_tally'] = 'first'
        
    agg = df.groupby(['invoice_no', 'gstin']).agg(agg_funcs).reset_index()
    
    if 'invoice_date' in agg.columns:
        agg['invoice_date'] = agg['invoice_date'].dt.strftime('%Y-%m-%d')
        
    return agg

def parse_gstr2b_file(file_path):
    file_path.seek(0)
    
    # Try to find B2B sheet, otherwise just use the first sheet
    xls = pd.ExcelFile(file_path)
    sheet_name = 'B2B' if 'B2B' in xls.sheet_names else xls.sheet_names[0]
    
    # Read first 20 rows to find header dynamically
    raw = pd.read_excel(file_path, sheet_name=sheet_name, nrows=20, header=None)
    header_row = None
    
    for i in range(raw.shape[0]):
        row_str = ' '.join([str(x).lower() for x in raw.iloc[i] if pd.notna(x)])
        if ('invoice number' in row_str or 'invoice no' in row_str) and 'gstin' in row_str:
            header_row = i
            break
            
    if header_row is None:
        # Fallback to row index 2 if not explicitly found
        header_row = 2
        
    file_path.seek(0)
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)
        
    col_map = {}
    for col in df.columns:
        col_lower = str(col).lower().strip().replace('\n', ' ')
        if 'invoice number' in col_lower or 'invoice no' in col_lower:
            col_map[col] = 'invoice_no'
        elif 'gstin' in col_lower:
            if 'gstin' not in col_map.values(): # Pick first matching to avoid overriding with multiple gstin cols
                col_map[col] = 'gstin'
        elif 'trade/legal name' in col_lower or 'trade name' in col_lower or 'legal name' in col_lower:
            col_map[col] = 'party_name_gst'
        elif 'invoice date' in col_lower:
            col_map[col] = 'invoice_date'
        elif 'taxable value' in col_lower:
            col_map[col] = 'taxable_value'
        elif 'integrated tax' in col_lower or col_lower == 'igst':
            col_map[col] = 'igst'
        elif 'central tax' in col_lower or col_lower == 'cgst':
            col_map[col] = 'cgst'
        elif 'state/ut tax' in col_lower or 'state tax' in col_lower or col_lower == 'sgst':
            col_map[col] = 'sgst'
        elif 'total tax' in col_lower or 'tax amount' in col_lower:
            col_map[col] = 'tax_amount_native'
            
    df.rename(columns=col_map, inplace=True)
    
    if 'invoice_no' not in df.columns:
        df['invoice_no'] = ''
    if 'gstin' not in df.columns:
        df['gstin'] = ''
        
    df['gstin'] = df['gstin'].astype(str).str.strip().str.upper()
    df['invoice_no'] = df['invoice_no'].astype(str).str.strip()
    
    for col in ['igst', 'cgst', 'sgst', 'taxable_value', 'tax_amount_native']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(',', '', regex=False).str.replace('₹', '', regex=False).str.replace('- 0', '0', regex=False).str.replace('-0', '0', regex=False).str.strip()
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        else:
            df[col] = 0.0
            
    # CRITICAL: Sum taxes
    if 'tax_amount_native' in df.columns and df['tax_amount_native'].sum() > 0:
        df['tax_amount'] = df['tax_amount_native']
    else:
        df['tax_amount'] = df['igst'] + df['cgst'] + df['sgst']
    
    if 'invoice_date' in df.columns:
        df['invoice_date'] = pd.to_datetime(df['invoice_date'], dayfirst=True, errors='coerce').dt.strftime('%Y-%m-%d')
        
    needed = ['gstin', 'invoice_no', 'invoice_date', 'taxable_value', 'tax_amount', 'party_name_gst']
    df = df[[c for c in needed if c in df.columns]]
    df = df.dropna(subset=['invoice_no'])
    df = df[df['invoice_no'] != 'nan']
    df = df[df['invoice_no'] != '']
    return df

def reconcile(file_tally, file_gstr2b):
    warnings = []
    
    # 1. Parse Files
    df_tally_raw, tally_raw_columns = parse_tally_file(file_tally)
    tally_df = aggregate_tally(df_tally_raw)
    gst_df = parse_gstr2b_file(file_gstr2b)
    
    if 'Bill No' not in tally_raw_columns and 'bill no' not in [c.lower() for c in tally_raw_columns]:
        warnings.append("⚠️ Tally file missing 'Bill No' column. Re-export from Tally with Bill No/Supplier Invoice No column for accurate matching.")
        
    # Convert dates to datetime for ± 3 days comparison later
    if 'invoice_date' in tally_df.columns:
        tally_df['invoice_date_dt'] = pd.to_datetime(tally_df['invoice_date'], errors='coerce')
    else:
        tally_df['invoice_date_dt'] = pd.NaT

    if 'invoice_date' in gst_df.columns:
        gst_df['invoice_date_dt'] = pd.to_datetime(gst_df['invoice_date'], errors='coerce')
    else:
        gst_df['invoice_date_dt'] = pd.NaT
        
    # Strategy 1: Exact Match on invoice_no
    merged = pd.merge(tally_df, gst_df, on='invoice_no', how='outer', indicator=True, suffixes=('_tally', '_gst'))
    
    matched = merged[merged['_merge'] == 'both'].copy()
    missing_in_gst = merged[merged['_merge'] == 'left_only'].copy()
    missing_in_tally = merged[merged['_merge'] == 'right_only'].copy()
    
    # Strategy 2: Fallback match by gstin + date (± 3 days) + taxable_value (± 1%)
    missing_in_gst_indices_to_drop = []
    missing_in_tally_indices_to_drop = []
    fallback_matches = []
    
    if not missing_in_gst.empty and not missing_in_tally.empty:
        for gst_idx, gst_row in missing_in_tally.iterrows():
            g_gstin = gst_row['gstin_gst']
            g_date = gst_row['invoice_date_dt_gst']
            g_val = gst_row['taxable_value_gst']
            
            candidates = missing_in_gst[
                (missing_in_gst['gstin_tally'] == g_gstin) &
                (~missing_in_gst.index.isin(missing_in_gst_indices_to_drop))
            ]
            
            if candidates.empty:
                continue
                
            if pd.notnull(g_date):
                candidates = candidates[
                    (candidates['invoice_date_dt_tally'].notnull()) &
                    (abs((candidates['invoice_date_dt_tally'] - g_date).dt.days) <= 3)
                ]
                
            if candidates.empty:
                continue
                
            # ± 1% tolerance
            candidates = candidates[
                abs(candidates['taxable_value_tally'] - g_val) <= (0.01 * abs(g_val) + 0.01)
            ]
            
            if not candidates.empty:
                t_idx = candidates.index[0]
                t_row = candidates.loc[t_idx]
                
                fallback_row = {
                    'invoice_no': t_row['invoice_no'],  # keep tally invoice no
                    'gstin_tally': t_row['gstin_tally'],
                    'invoice_date_tally': t_row['invoice_date_tally'],
                    'taxable_value_tally': t_row['taxable_value_tally'],
                    'tax_amount_tally': t_row['tax_amount_tally'],
                    'party_name_tally': t_row.get('party_name_tally'),
                    'gstin_gst': g_gstin,
                    'invoice_date_gst': gst_row['invoice_date_gst'],
                    'taxable_value_gst': g_val,
                    'tax_amount_gst': gst_row['tax_amount_gst'],
                    'invoice_no_gst': gst_row['invoice_no'],
                    'party_name_gst': gst_row.get('party_name_gst'),
                    '_merge': 'fallback_match'
                }
                fallback_matches.append(fallback_row)
                
                missing_in_gst_indices_to_drop.append(t_idx)
                missing_in_tally_indices_to_drop.append(gst_idx)
                
    missing_in_gst = missing_in_gst.drop(index=missing_in_gst_indices_to_drop)
    missing_in_tally = missing_in_tally.drop(index=missing_in_tally_indices_to_drop)
    
    if fallback_matches:
        matched = pd.concat([matched, pd.DataFrame(fallback_matches)], ignore_index=True)
        
    final_matched = []
    final_mismatched = []
    
    for _, row in matched.iterrows():
        t_val = row.get('taxable_value_tally', 0)
        g_val = row.get('taxable_value_gst', 0)
        t_tax = row.get('tax_amount_tally', 0)
        g_tax = row.get('tax_amount_gst', 0)
        
        if pd.isna(t_val): t_val = 0
        if pd.isna(g_val): g_val = 0
        if pd.isna(t_tax): t_tax = 0
        if pd.isna(g_tax): g_tax = 0
        
        val_diff = abs(t_val - g_val)
        tax_diff = abs(t_tax - g_tax)
        
        mismatches = []
        if val_diff > 1.0:
            mismatches.append(f"Taxable Value (GST: {g_val}, Tally: {t_val})")
        if tax_diff > 1.0:
            mismatches.append(f"Tax Amount (GST: {g_tax}, Tally: {t_tax})")
            
        row_dict = row.to_dict()
        if not mismatches:
            row_dict['Status'] = 'Matched'
            if val_diff == 0 and tax_diff == 0:
                row_dict['Reason'] = 'Exact Match' if row.get('_merge') == 'both' else 'Fallback Match'
            else:
                row_dict['Reason'] = 'Matched (Rounding Difference)'
            final_matched.append(row_dict)
        else:
            row_dict['Status'] = 'Mismatched'
            row_dict['Reason'] = " | ".join(mismatches)
            final_mismatched.append(row_dict)
            
    df_matched = pd.DataFrame(final_matched)
    df_mismatched = pd.DataFrame(final_mismatched)
    
    missing_in_gst['Status'] = 'Missing in GST'
    missing_in_gst['Reason'] = ''
    missing_in_tally['Status'] = 'Missing in Tally'
    missing_in_tally['Reason'] = ''
    
    # Clean up datetimes and merge cols
    for df in [df_matched, df_mismatched, missing_in_gst, missing_in_tally]:
        if not df.empty:
            for col in ['invoice_date_dt_tally', 'invoice_date_dt_gst', 'invoice_date_dt', '_merge']:
                if col in df.columns:
                    df.drop(columns=[col], inplace=True)
                    
    all_dfs = [df for df in [df_matched, df_mismatched, missing_in_gst, missing_in_tally] if not df.empty]
    merged_output = pd.concat(all_dfs, ignore_index=True) if all_dfs else pd.DataFrame()
    
    summary_metrics = {
        'Total Records': len(merged_output),
        'Matched': len(df_matched),
        'Mismatched': len(df_mismatched),
        'Missing in GST': len(missing_in_gst),
        'Missing in Tally': len(missing_in_tally),
    }
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        if not merged_output.empty:
            merged_output.to_excel(writer, sheet_name='Summary', index=False)
            
        for name, df in [('Matched', df_matched), ('Mismatched', df_mismatched), 
                         ('Missing in GST', missing_in_gst), ('Missing in Tally', missing_in_tally)]:
            if not df.empty:
                df_to_save = df.copy()
                if name == 'Missing in GST' and 'party_name_gst' in df_to_save.columns:
                    df_to_save = df_to_save.drop(columns=['party_name_gst'])
                elif name == 'Missing in Tally' and 'party_name_tally' in df_to_save.columns:
                    df_to_save = df_to_save.drop(columns=['party_name_tally'])
                df_to_save.to_excel(writer, sheet_name=name[:31], index=False)
                
        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            df_to_format = merged_output if sheet_name == 'Summary' else next((mapping for key, mapping in {'Matched': df_matched, 'Mismatched': df_mismatched, 'Missing in GST': missing_in_gst, 'Missing in Tally': missing_in_tally}.items() if key[:31] == sheet_name), pd.DataFrame())
            if not df_to_format.empty:
                for i, col in enumerate(df_to_format.columns):
                    safe_series = df_to_format[col].fillna('').astype(str)
                    col_max_len = safe_series.map(len).max()
                    col_max_len = 0 if pd.isna(col_max_len) else col_max_len
                    worksheet.set_column(i, i, int(max(col_max_len, len(str(col)))) + 2)
                    
    output.seek(0)
    return summary_metrics, output, warnings, missing_in_gst

def generate_supplier_defaulters_list(missing_in_gst_df):
    """
    Generate a professional Supplier Defaulters List for emails/follow-ups
    """
    if missing_in_gst_df.empty:
        return None
    
    # Select and rename columns for the defaulter list
    defaulter_cols = {
        'invoice_no': 'Voucher No (Tally)',
        'gstin_tally': 'Supplier GSTIN',
        'party_name_tally': 'Supplier Name',
        'invoice_date_tally': 'Invoice Date',
        'taxable_value_tally': 'Taxable Value (₹)',
        'tax_amount_tally': 'Tax Amount (₹)',
    }
    
    # Create the defaulter dataframe
    defaulters = missing_in_gst_df[[col for col in defaulter_cols.keys() if col in missing_in_gst_df.columns]].copy()
    defaulters.rename(columns=defaulter_cols, inplace=True)
    
    # Calculate Total Invoice Value
    if 'Taxable Value (₹)' in defaulters.columns and 'Tax Amount (₹)' in defaulters.columns:
        defaulters['Total Invoice Value (₹)'] = (
            defaulters['Taxable Value (₹)'].fillna(0) + 
            defaulters['Tax Amount (₹)'].fillna(0)
        )
    
    # Add Action Required column
    defaulters['Action Required'] = 'Supplier must file GSTR-1'
    
    # Sort by Supplier Name then by Date
    if 'Supplier Name' in defaulters.columns:
        defaulters.sort_values(by=['Supplier Name', 'Invoice Date'], inplace=True)
    
    # Reset index
    defaulters.reset_index(drop=True, inplace=True)
    
    # Add summary row at the top
    summary_data = {
        'Voucher No (Tally)': 'SUMMARY',
        'Supplier GSTIN': '',
        'Supplier Name': f'Total Missing Invoices: {len(defaulters)}',
        'Invoice Date': '',
        'Taxable Value (₹)': defaulters['Taxable Value (₹)'].sum(),
        'Tax Amount (₹)': defaulters['Tax Amount (₹)'].sum(),
        'Total Invoice Value (₹)': defaulters['Total Invoice Value (₹)'].sum(),
        'Action Required': ''
    }
    
    # Generate Excel with formatting
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        defaulters.to_excel(writer, sheet_name='Supplier Defaulters', index=False, startrow=2)
        
        # Get the workbook and worksheet
        workbook = writer.book
        worksheet = writer.sheets['Supplier Defaulters']
        
        # Define formats
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#4472C4',
            'font_color': 'white',
            'border': 1,
            'text_wrap': True
        })
        
        summary_format = workbook.add_format({
            'bold': True,
            'bg_color': '#FFC000',
            'border': 1
        })
        
        title_format = workbook.add_format({
            'bold': True,
            'font_size': 14,
            'font_color': '#1F4E78'
        })
        
        # Write title
        worksheet.write(0, 0, 'SUPPLIER DEFAULTERS LIST - GSTR-1 NOT FILED', title_format)
        
        # Write summary row
        for col_idx, (col_name, value) in enumerate(summary_data.items()):
            worksheet.write(2, col_idx, value, summary_format if col_idx >= 4 else summary_format)
        
        # Apply header format
        for col_idx, col_name in enumerate(defaulters.columns):
            worksheet.write(3, col_idx, col_name, header_format)
        
        # Set column widths
        column_widths = {
            'Voucher No (Tally)': 15,
            'Supplier GSTIN': 20,
            'Supplier Name': 35,
            'Invoice Date': 12,
            'Taxable Value (₹)': 18,
            'Tax Amount (₹)': 15,
            'Total Invoice Value (₹)': 20,
            'Action Required': 30
        }
        
        for col_idx, col_name in enumerate(defaulters.columns):
            width = column_widths.get(col_name, 15)
            worksheet.set_column(col_idx, col_idx, width)
        
        # Add number format for currency columns
        money_format = workbook.add_format({'num_format': '#,##0.00'})
        for col_idx, col_name in enumerate(defaulters.columns):
            if 'Value' in col_name or 'Amount' in col_name:
                worksheet.set_column(col_idx, col_idx, None, money_format)
    
    output.seek(0)
    return output