import streamlit as st
import pandas as pd
import gspread
import json

# --- CONFIGURATION ---
# PASTE YOUR GOOGLE SHEET ID HERE (Keep the quotes!)
GOOGLE_SHEET_ID = "PASTE_YOUR_LONG_GOOGLE_SHEET_ID_HERE"
MAIN_TAB_NAME   = "Processed Data"
DUP_TAB_NAME    = "Duplicates Found"
# ---------------------

def get_google_client(uploaded_file, pasted_text):
    """Tries to get Google Client from File Upload OR Pasted Text."""
    creds = None
    
    # Priority 1: File Upload
    if uploaded_file is not None:
        try:
            creds = json.load(uploaded_file)
        except Exception as e:
            st.error(f"❌ Error reading uploaded file: {e}")
            return None

    # Priority 2: Pasted Text
    elif pasted_text.strip():
        try:
            creds = json.loads(pasted_text, strict=False)
        except Exception as e:
            st.error(f"❌ Error reading pasted text. Make sure you copied the WHOLE content.\nDetails: {e}")
            return None

    # Connect if we have creds
    if creds:
        try:
            return gspread.service_account_from_dict(creds)
        except Exception as e:
            st.error(f"❌ Login failed: {e}")
            return None
            
    return None

def main():
    st.set_page_config(page_title="Excel Merger", page_icon="📂")
    st.title("📂 Excel Merger & Google Uploader")

    # --- SIDEBAR: AUTHENTICATION ---
    st.sidebar.header("🔑 Authentication")
    
    # Option 1: File
    uploaded_key = st.sidebar.file_uploader("Option 1: Upload JSON File", type=['json'])
    
    st.sidebar.markdown("--- OR ---")
    
    # Option 2: Paste Text
    pasted_key = st.sidebar.text_area("Option 2: Paste JSON Content Here", height=200)

    # --- MAIN CONTENT ---
    uploaded_files = st.file_uploader("Upload Excel Files to Merge", type=['xlsx'], accept_multiple_files=True)

    if st.button("Merge & Upload"):
        # 1. CONNECT TO GOOGLE
        gc = get_google_client(uploaded_key, pasted_key)
        
        if not gc:
            st.error("❌ You must upload a JSON file OR paste the key text in the sidebar!")
            st.stop()

        try:
            sh = gc.open_by_key(GOOGLE_SHEET_ID)
        except Exception as e:
            st.error(f"❌ Connected to Google, but could not open Sheet.\nCheck your Sheet ID: {GOOGLE_SHEET_ID}\nError: {e}")
            st.stop()

        if not uploaded_files:
            st.warning("Please upload at least one Excel file.")
            return

        st.info("Reading files...")
        
        # 2. PROCESS FILES
        all_dfs = []
        for file in uploaded_files:
            try:
                # header=1 means Row 2 is the header
                df = pd.read_excel(file, header=1)
                all_dfs.append(df)
            except Exception as e:
                st.error(f"Error reading {file.name}: {e}")
                return

        if not all_dfs:
            st.error("No valid data found.")
            return

        # Merge
        master_df = pd.concat(all_dfs, ignore_index=True)
        st.write(f"✅ Merged {len(master_df)} rows.")

        # Logic
        if 'Transaction ID' in master_df.columns:
            master_df['Date'] = master_df['Transaction ID'].apply(
                lambda x: f"{x.split('-')[1][:2]}/{x.split('-')[1][2:4]}/{x.split('-')[1][4:8]}" 
                if isinstance(x, str) and '-' in x else None
            )

        if 'Mobile Number' in master_df.columns:
            master_df['Mobile Number'] = master_df['Mobile Number'].astype(str).str.replace(r'\.0$', '', regex=True)

        # Duplicates
        dup_cols = ['Name', 'Mobile Number']
        if set(dup_cols).issubset(master_df.columns):
            dup_mask = master_df.duplicated(subset=dup_cols, keep=False)
            df_dupes = master_df[dup_mask][dup_cols].drop_duplicates()
        else:
            df_dupes = pd.DataFrame()

        # 3. UPLOAD
        def upload(tab, df):
            try:
                ws = sh.worksheet(tab)
                ws.clear()
            except gspread.WorksheetNotFound:
                ws = sh.add_worksheet(title=tab, rows=len(df)+50, cols=len(df.columns))
            
            df = df.fillna('')
            ws.update([df.columns.values.tolist()] + df.values.tolist())
            st.success(f"✅ Uploaded to '{tab}'")

        try:
            upload(MAIN_TAB_NAME, master_df)
            if not df_dupes.empty:
                upload(DUP_TAB_NAME, df_dupes)
            else:
                st.info("No duplicates found.")
        except Exception as e:
            st.error(f"Upload failed: {e}")

if __name__ == "__main__":
    main()
