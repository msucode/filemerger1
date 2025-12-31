import pandas as pd
import gspread
import sys
import os
import json

# --- CONFIGURATION ---
EXCEL_FILE_NAME = "input_data.xlsx"       # Ensure this file exists in your repo
GOOGLE_SHEET_ID = "PASTE_YOUR_SHEET_ID_HERE"
MAIN_TAB_NAME   = "All Data"
DUP_TAB_NAME    = "Duplicate Entries"
# ---------------------

def main():
    # 1. AUTHENTICATE USING GITHUB SECRET
    creds_json = os.environ.get("GOOGLE_CREDENTIALS")
    if not creds_json:
        print("CRITICAL ERROR: 'GOOGLE_CREDENTIALS' secret is missing in GitHub Settings.")
        sys.exit(1)
    
    # Save secret to a temporary file for gspread to read
    with open("temp_creds.json", "w") as f:
        f.write(creds_json)

    # 2. READ EXCEL
    try:
        print(f"Reading {EXCEL_FILE_NAME}...")
        df = pd.read_excel(EXCEL_FILE_NAME, header=1) # header=1 means Row 2 is header
    except FileNotFoundError:
        print(f"ERROR: {EXCEL_FILE_NAME} not found. Did you upload it?")
        sys.exit(1)

    # 3. FIND DUPLICATES (Name + Mobile Number)
    print("Processing duplicates...")
    dupe_mask = df.duplicated(subset=['Name', 'Mobile Number'], keep=False)
    df_dupes = df[dupe_mask][['Name', 'Mobile Number']]

    # 4. UPLOAD TO GOOGLE
    print("Connecting to Google Sheets...")
    gc = gspread.service_account(filename="temp_creds.json")
    sh = gc.open_by_key(GOOGLE_SHEET_ID)

    def upload(tab_name, data):
        try:
            ws = sh.worksheet(tab_name)
            ws.clear()
        except gspread.WorksheetNotFound:
            ws = sh.add_worksheet(title=tab_name, rows=len(data)+20, cols=len(data.columns))
        
        # Convert to list and upload
        data = data.fillna('')
        ws.update([data.columns.values.tolist()] + data.values.tolist())
        print(f"Uploaded to '{tab_name}'")

    upload(MAIN_TAB_NAME, df)
    upload(DUP_TAB_NAME, df_dupes)

    # Cleanup
    os.remove("temp_creds.json")
    print("Success.")

if __name__ == "__main__":
    main()
