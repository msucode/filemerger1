import pandas as pd
import gspread
import sys
import os
import json
from excel_reader import read_excel

# ================= CONFIGURATION =================
# 1. PASTE YOUR GOOGLE SHEET ID BELOW (Keep the quotes!)
GOOGLE_SHEET_ID = "1Tlq1386a67nifXnx4_l9rlX1wbWyylJgQrdiZsOEIxY"

# 2. FILE SETTINGS
EXCEL_FILE_NAME = "input_data.xlsx"
MAIN_TAB_NAME   = "Processed Data"
DUP_TAB_NAME    = "Duplicates Found"
# =================================================

def extract_date(transaction_id):
    """Preserved logic: Extract date from Transaction ID (DDMMYYYY)"""
    if isinstance(transaction_id, str):
        parts = transaction_id.split('-')
        if len(parts) >= 2 and len(parts[1]) >= 8:
            date_part = parts[1][:8]  # Extract first 8 chars of the second part
            return f"{date_part[:2]}/{date_part[2:4]}/{date_part[4:]}"
    return None

def main():
    print("--- STARTING ROBOT ---")

    # 1. AUTHENTICATE (Using the Secret from GitHub Settings)
    creds_json = os.environ.get("GOOGLE_CREDENTIALS")
    if not creds_json:
        print("CRITICAL ERROR: 'GOOGLE_CREDENTIALS' secret is missing in GitHub Settings.")
        sys.exit(1)
    
    # Write secret to temp file for gspread
    with open("temp_creds.json", "w") as f:
        f.write(creds_json)

    # 2. READ EXCEL (Uses your excel_reader.py logic)
    print(f"Reading {EXCEL_FILE_NAME}...")
    try:
        # Note: excel_reader.py uses header=1 (skips row 1)
        df = pd.read_excel(EXCEL_FILE_NAME, header=1)
    except FileNotFoundError:
        print(f"ERROR: {EXCEL_FILE_NAME} not found. Did you upload it to the repo?")
        sys.exit(1)

    # 3. APPLY YOUR LOGIC (Extract Date)
    print("Applying data logic...")
    if 'Transaction ID' in df.columns:
        df['Date'] = df['Transaction ID'].apply(extract_date)
    else:
        print("Warning: 'Transaction ID' column not found. Skipping date extraction.")

    # 4. IDENTIFY DUPLICATES (Name + Mobile)
    print("Finding duplicates...")
    # Clean mobile numbers to ensure accurate matching (remove .0 if exists)
    if 'Mobile Number' in df.columns:
        df['Mobile Number'] = df['Mobile Number'].astype(str).str.replace('.0', '', regex=False)

    dup_cols = ['Name', 'Mobile Number']
    # Check if columns exist before looking for duplicates
    if set(dup_cols).issubset(df.columns):
        duplicate_mask = df.duplicated(subset=dup_cols, keep=False)
        df_dupes = df[duplicate_mask][dup_cols].drop_duplicates()
    else:
        print(f"Warning: Columns {dup_cols} not found. Skipping duplicate check.")
        df_dupes = pd.DataFrame() # Empty

    # 5. UPLOAD TO GOOGLE SHEETS
    print("Connecting to Google Sheets...")
    try:
        gc = gspread.service_account(filename="temp_creds.json")
        sh = gc.open_by_key(GOOGLE_SHEET_ID)
    except Exception as e:
        print(f"ERROR: Connection failed. Check your Sheet ID and Share settings. Details: {e}")
        sys.exit(1)

    def upload(tab_name, data):
        try:
            ws = sh.worksheet(tab_name)
            ws.clear()
        except gspread.WorksheetNotFound:
            ws = sh.add_worksheet(title=tab_name, rows=len(data)+50, cols=len(data.columns))
        
        # Google Sheets needs NaNs replaced by empty strings
        data = data.fillna('')
        ws.update([data.columns.values.tolist()] + data.values.tolist())
        print(f"Uploaded {len(data)} rows to tab: '{tab_name}'")

    upload(MAIN_TAB_NAME, df)
    if not df_dupes.empty:
        upload(DUP_TAB_NAME, df_dupes)
    else:
        print("No duplicates found to upload.")

    # Cleanup
    if os.path.exists("temp_creds.json"):
        os.remove("temp_creds.json")
    print("--- SUCCESS ---")

if __name__ == "__main__":
    main()
