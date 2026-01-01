import streamlit as st
import pandas as pd
import gspread
import json

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

    # --- SIDEBAR: CONFIGURATION ---
    st.sidebar.header("⚙️ Configuration")
    
    # 1. GOOGLE SHEET ID INPUT
    st.sidebar.subheader("1. Destination Sheet")
    sheet_id = st.sidebar.text_input("1Tlq1386a67nifXnx4_l9rlX1wbWyylJgQrdiZsOEIxY", help="https://docs.google.com/spreadsheets/d/1Tlq1386a67nifXnx4_l9rlX1wbWyylJgQrdiZsOEIxY/edit?gid=0#gid=0
")

    # 2. AUTHENTICATION INPUTS
    st.sidebar.subheader("2. Authentication Key")
    st.sidebar.info("Upload your 'client_secret.json' OR paste the text.")
    uploaded_key = st.sidebar.file_uploader("Upload JSON File", type=['json'])
    st.sidebar.text("--- OR ---")
    pasted_key = st.sidebar.text_area("Paste JSON Content", height=150, help="Open your json file in Notepad, copy all, paste here.")

    # --- MAIN CONTENT ---
    st.subheader("3. Upload Excel Files")
    uploaded_files = st.file_uploader("Drop Excel files here to merge", type=['xlsx'], accept_multiple_files=True)

    if st.button("Merge & Upload"):
        # VALIDATION
        if not sheet_id:
            st.error("❌ Please paste your Google Sheet ID in the Sidebar.")
            st.stop()
        
        if not uploaded_files:
            st.warning("⚠️ Please upload at least one Excel file.")
            return

        # 1. CONNECT TO GOOGLE
        gc = get_google_client(uploaded_key, pasted_key)
        
        if not gc:
            st.error("❌ Authentication Missing! Upload your JSON file or paste the key in the sidebar.")
            st.stop()

        try:
            sh = gc.open_by_key(sheet_id)
        except Exception as e:
            st.error(f"❌ Connection Successful, but could not find the Sheet.\n\n1. Check the ID: {sheet_id}\n2. Did you SHARE the sheet with the email inside your JSON key?\n\nError: {e}")
            st.stop()

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
            upload("Processed Data", master_df)
            if not df_dupes.empty:
                upload("Duplicates Found", df_dupes)
            else:
                st.info("No duplicates found.")
        except Exception as e:
            st.error(f"Upload failed: {e}")

if __name__ == "__main__":
    main()
    
