import streamlit as st
import pandas as pd
import gspread
import json

# --- HELPER FUNCTIONS ---
def get_google_client(uploaded_file, pasted_text):
    """Connects to Google Sheets using the uploaded file or pasted text."""
    creds = None
    
    # Method 1: File Upload
    if uploaded_file is not None:
        try:
            creds = json.load(uploaded_file)
        except Exception as e:
            st.error(f"Error reading file: {e}")
            return None

    # Method 2: Pasted Text
    elif pasted_text.strip():
        try:
            # strict=False allows control characters like newlines
            creds = json.loads(pasted_text, strict=False)
        except Exception as e:
            st.error(f"Error reading text: {e}")
            return None

    # Connect
    if creds:
        try:
            return gspread.service_account_from_dict(creds)
        except Exception as e:
            st.error(f"Login failed: {e}")
            return None
    return None

def extract_date_safe(tid):
    """Safely extracts date from Transaction ID (e.g., TXN-01012023-001)."""
    if not isinstance(tid, str) or '-' not in tid:
        return None
    
    try:
        # Split by dash and take the second part (the date)
        parts = tid.split('-')
        if len(parts) < 2:
            return None
            
        date_part = parts[1]
        # Ensure it looks like a date (DDMMYYYY)
        if len(date_part) >= 8:
            return f"{date_part[:2]}/{date_part[2:4]}/{date_part[4:8]}"
    except:
        return None
    return None

# --- MAIN APP ---
def main():
    st.set_page_config(page_title="Excel Merger", page_icon="📂")
    st.title("📂 Excel Merger & Google Uploader")

    # --- SIDEBAR ---
    st.sidebar.header("Configuration")
    
    # 1. Sheet ID
    st.sidebar.subheader("1. Destination Sheet")
    sheet_id = st.sidebar.text_input("Paste Google Sheet ID:")

    # 2. Auth
    st.sidebar.subheader("2. Authentication")
    st.sidebar.info("Upload 'client_secret.json' OR paste content.")
    
    uploaded_key = st.sidebar.file_uploader("Upload Key File", type=['json'])
    st.sidebar.text("--- OR ---")
    pasted_key = st.sidebar.text_area("Paste Key Content", height=150)

    # --- MAIN BODY ---
    st.subheader("3. Upload Excel Files")
    uploaded_files = st.file_uploader("Drop files here", type=['xlsx'], accept_multiple_files=True)

    if st.button("Merge & Upload"):
        # Checks
        if not sheet_id:
            st.error("❌ Missing Google Sheet ID in sidebar.")
            st.stop()
        
        gc = get_google_client(uploaded_key, pasted_key)
        if not gc:
            st.error("❌ Missing Authentication (Upload JSON or Paste Key).")
            st.stop()

        if not uploaded_files:
            st.warning("⚠️ Please upload at least one Excel file.")
            return

        # 1. Connect
        try:
            sh = gc.open_by_key(sheet_id)
        except Exception as e:
            st.error(f"❌ Connection error. Check Sheet ID. Error: {e}")
            st.stop()

        st.info("Processing files...")

        # 2. Merge
        all_dfs = []
        for file in uploaded_files:
            try:
                # header=1 means Row 2 is header
                df = pd.read_excel(file, header=1)
                all_dfs.append(df)
            except Exception as e:
                st.error(f"Error reading {file.name}: {e}")
                return

        if not all_dfs:
            st.error("No valid data found.")
            return

        master_df = pd.concat(all_dfs, ignore_index=True)
        st.write(f"✅ Merged {len(master_df)} rows.")

        # 3. Clean Data
        if 'Transaction ID' in master_df.columns:
            master_df['Date'] = master_df['Transaction ID'].apply(extract_date_safe)

        if 'Mobile Number' in master_df.columns:
            # Remove .0 from phone numbers
            master_df['Mobile Number'] = master_df['Mobile Number'].astype(str).str.replace(r'\.0$', '', regex=True)

        # 4. Find Duplicates
        dup_cols = ['Name', 'Mobile Number']
        if set(dup_cols).issubset(master_df.columns):
            dup_mask = master_df.duplicated(subset=dup_cols, keep=False)
            df_dupes = master_df[dup_mask][dup_cols].drop_duplicates()
        else:
            df_dupes = pd.DataFrame()

        # 5. Upload Helper
        def upload_to_tab(tab_name, data):
            try:
                try:
                    ws = sh.worksheet(tab_name)
                    ws.clear()
                except gspread.WorksheetNotFound:
                    ws = sh.add_worksheet(title=tab_name, rows=len(data)+50, cols=len(data.columns))
                
                # Fill NaN with empty string for Google Sheets
                clean_data = data.fillna('')
                ws.update([clean_data.columns.values.tolist()] + clean_data.values.tolist())
                st.success(f"✅ Uploaded to '{tab_name}'")
            except Exception as e:
                st.error(f"Failed to upload {tab_name}: {e}")

        # Upload
        upload_to_tab("Processed Data", master_df)
        
        if not df_dupes.empty:
            upload_to_tab("Duplicates Found", df_dupes)
        else:
            st.info("No duplicates found.")

if __name__ == "__main__":
    main()
