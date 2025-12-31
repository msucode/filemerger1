import streamlit as st
import pandas as pd
import gspread
import json

# --- CONFIGURATION ---
# PASTE YOUR GOOGLE SHEET ID HERE (Keep the quotes!)
GOOGLE_SHEET_ID = "1Tlq1386a67nifXnx4_l9rlX1wbWyylJgQrdiZsOEIxY"
MAIN_TAB_NAME   = "Processed Data"
DUP_TAB_NAME    = "Duplicates Found"
# ---------------------

def connect_to_google():
    """Connects to Google Sheets using Streamlit Secrets."""
    try:
        # Load credentials from Streamlit secrets
        creds_dict = json.loads(st.secrets["GOOGLE_CREDENTIALS"])
        gc = gspread.service_account_from_dict(creds_dict)
        return gc.open_by_key(GOOGLE_SHEET_ID)
    except Exception as e:
        st.error(f"Google Connection Error: {e}")
        return None

def extract_date(transaction_id):
    """Extracts date from Transaction ID (DDMMYYYY logic)."""
    if isinstance(transaction_id, str):
        parts = transaction_id.split('-')
        if len(parts) >= 2 and len(parts[1]) >= 8:
            date_part = parts[1][:8]
            return f"{date_part[:2]}/{date_part[2:4]}/{date_part[4:]}"
    return None

def main():
    st.title("📂 Excel Merger & Google Uploader")

    # 1. FILE UPLOAD WIDGET
    uploaded_files = st.file_uploader("Upload Excel Files", type=['xlsx'], accept_multiple_files=True)

    # 2. SUBMIT BUTTON
    if st.button("Merge & Upload"):
        if not uploaded_files:
            st.warning("Please upload at least one file.")
            return

        st.info("Reading and merging files...")
        
        # Merge Files
        all_dfs = []
        for file in uploaded_files:
            try:
                # header=1 means Row 2 is the header (Skip Row 1)
                df = pd.read_excel(file, header=1)
                all_dfs.append(df)
            except Exception as e:
                st.error(f"Error reading {file.name}: {e}")
                return

        if not all_dfs:
            st.error("No data found.")
            return

        # Combine into one Master DataFrame
        master_df = pd.concat(all_dfs, ignore_index=True)
        st.write(f"✅ Merged {len(master_df)} total rows.")

        # Data Logic: Extract Date
        if 'Transaction ID' in master_df.columns:
            master_df['Date'] = master_df['Transaction ID'].apply(extract_date)

        # Clean Mobile Numbers (remove .0)
        if 'Mobile Number' in master_df.columns:
            master_df['Mobile Number'] = master_df['Mobile Number'].astype(str).str.replace(r'\.0$', '', regex=True)

        # Duplicate Logic
        dup_cols = ['Name', 'Mobile Number']
        if set(dup_cols).issubset(master_df.columns):
            dup_mask = master_df.duplicated(subset=dup_cols, keep=False)
            df_dupes = master_df[dup_mask][dup_cols].drop_duplicates()
        else:
            df_dupes = pd.DataFrame()

        # Google Upload
        sh = connect_to_google()
        if sh:
            try:
                # Upload Main Data
                ws_main = get_or_create_worksheet(sh, MAIN_TAB_NAME, master_df)
                set_worksheet_data(ws_main, master_df)
                st.success(f"Uploaded main data to '{MAIN_TAB_NAME}'")

                # Upload Duplicates
                if not df_dupes.empty:
                    ws_dup = get_or_create_worksheet(sh, DUP_TAB_NAME, df_dupes)
                    set_worksheet_data(ws_dup, df_dupes)
                    st.success(f"Uploaded duplicates to '{DUP_TAB_NAME}'")
                else:
                    st.info("No duplicates found.")
                    
            except Exception as e:
                st.error(f"Upload failed: {e}")

def get_or_create_worksheet(sh, title, df):
    try:
        return sh.worksheet(title)
    except gspread.WorksheetNotFound:
        return sh.add_worksheet(title=title, rows=len(df)+50, cols=len(df.columns))

def set_worksheet_data(ws, df):
    ws.clear()
    df_clean = df.fillna('')
    ws.update([df_clean.columns.values.tolist()] + df_clean.values.tolist())

if __name__ == "__main__":
    main()
