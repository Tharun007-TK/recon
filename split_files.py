import pandas as pd
import os

print("🚀 Starting split process...")
file_path = '2A Reconciliation.xlsx'

if not os.path.exists(file_path):
    print(f" ERROR: '{file_path}' not found in {os.getcwd()}")
else:
    # 1. Read the first 10 rows WITHOUT setting a header to find the real header row
    raw_df = pd.read_excel(file_path, header=None, nrows=10)
    
    print("\n Scanning first 5 rows to find headers...")
    print(raw_df.head())
    
    header_row_idx = None
    for i in range(raw_df.shape[0]):
        # Convert row to string and look for keywords
        row_str = ' '.join([str(x).lower() for x in raw_df.iloc[i] if pd.notna(x)])
        if 'invoice' in row_str or 'gstin' in row_str or 'taxable' in row_str:
            header_row_idx = i
            print(f"\n✅ Found real header row at index {i}!")
            print(f"Headers: {raw_df.iloc[i].tolist()}")
            break
            
    if header_row_idx is None:
        print("\n❌ Could not find header row. Please check the file structure manually.")
    else:
        # 2. Read the file again, skipping the top rows and using the found row as header
        df = pd.read_excel(file_path, header=header_row_idx)
        print(f"\n📊 Total columns found: {len(df.columns)}")
        print(f"📋 Actual Columns: {df.columns.tolist()}")

        # 3. Define the columns we want to extract
        gst_cols = ['GSTIN', 'Invoice Number', 'Invoice Date', 'Taxable Value', 'IGST', 'CGST', 'SGST']
        tally_cols = ['Party Name', 'GSTIN', 'Invoice Number', 'Invoice Date', 'Taxable Value', 'IGST', 'CGST', 'SGST']

        # Helper function to find columns safely (case-insensitive)
        def extract_cols(df, target_cols):
            new_df = pd.DataFrame()
            for col in target_cols:
                matched = [c for c in df.columns if col.lower() in str(c).lower()]
                if matched:
                    new_df[col] = df[matched[0]]
                else:
                    print(f"⚠️ Warning: Could not find column '{col}' in the file.")
            return new_df

        # 4. Create the DataFrames
        gst_df = extract_cols(df, gst_cols)
        tally_df = extract_cols(df, tally_cols)

        # 5. Save to new Excel files in the data folder
        data_dir = 'data'
        if not os.path.exists(data_dir):
            os.makedirs(data_dir)
            
        gst_path = os.path.join(data_dir, 'GST_Source.xlsx')
        tally_path = os.path.join(data_dir, 'Tally_Source.xlsx')
        
        gst_df.to_excel(gst_path, index=False)
        tally_df.to_excel(tally_path, index=False)

        print(f"\n🎉 SUCCESS!")
        print(f"✅ Created: {gst_path} ({len(gst_df)} records)")
        print(f"✅ Created: {tally_path} ({len(tally_df)} records)")
        print(f"👉 Check your '{data_dir}' folder for the new files!")