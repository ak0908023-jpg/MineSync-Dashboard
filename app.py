import streamlit as st
import pandas as pd
import datetime
import io
from fpdf import FPDF
import plotly.express as px

# -----------------------------------------------------------------------------
# 1. Dashboard Configuration & Custom Styling
# -----------------------------------------------------------------------------
st.set_page_config(page_title="MineSync Compliance Manager", page_icon="⚙️", layout="wide")

st.markdown("""
    <style>
        /* Main background and fonts */
        .main { background-color: #f4f6f9; }
        h1, h2, h3 { color: #1e3d59; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
        
        /* Stylish Metric Cards */
        div[data-testid="metric-container"] {
            background: linear-gradient(135deg, #ffffff 0%, #f8f9fa 100%);
            border-left: 5px solid #1e3d59;
            border: 1px solid #e0e0e0;
            padding: 20px; 
            border-radius: 12px; 
            box-shadow: 0 4px 15px rgba(0,0,0,0.05);
            transition: transform 0.2s ease;
        }
        div[data-testid="metric-container"]:hover {
            transform: translateY(-5px);
            box-shadow: 0 8px 25px rgba(0,0,0,0.1);
        }
        
        /* Big Navigation Buttons */
        .stButton>button {
            background: linear-gradient(to right, #1e3d59, #2b577d); 
            color: white; 
            border-radius: 8px; 
            border: none;
            box-shadow: 0 4px 10px rgba(0,0,0,0.15); 
            transition: all 0.3s ease;
            height: 60px; 
            font-size: 16px; 
            font-weight: bold;
        }
        .stButton>button:hover { 
            background: linear-gradient(to right, #2b577d, #3b6998); 
            transform: translateY(-2px); 
        }
        
        /* Logout Button Override */
        div[data-testid="stSidebar"] .stButton>button { 
            background: #ff4b4b; height: auto; margin-top: 10px;
        }
        
        /* Search Bar Highlighting */
        .stTextInput>div>div>input {
            border: 2px solid #1e3d59 !important;
            border-radius: 8px;
        }
        
        /* ----------------------------------- */
        /* CUSTOM TAB COLORS                   */
        /* ----------------------------------- */
        div[data-testid="stTabs"] button[data-baseweb="tab"] {
            background-color: #e2e8f0; 
            color: #1e3d59; 
            border-radius: 8px 8px 0px 0px; 
            padding: 10px 20px;
            font-weight: bold;
            border: 1px solid #cbd5e1;
            border-bottom: none;
            margin-right: 5px;
            transition: all 0.3s ease;
        }
        div[data-testid="stTabs"] button[data-baseweb="tab"]:hover {
            background-color: #cbd5e1;
        }
        div[data-testid="stTabs"] button[aria-selected="true"] {
            background: linear-gradient(to right, #1e3d59, #2b577d) !important;
            color: white !important;
            border: none;
        }
        div[data-testid="stTabs"] div[data-baseweb="tab-highlight"] {
            display: none;
        }
    </style>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# 1.5 Security & Authentication Module
# -----------------------------------------------------------------------------
USER_CREDENTIALS = {
    "admin": "nalco2026",
    "shift_manager": "panchpatmali123",
    "safety_officer": "dgms2026"
}

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.markdown("<br><br><h1 style='text-align: center; color: #1e3d59;'>🔒 MineSync Secure Portal</h1>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.info("Enter your administrative credentials to access the DGMS dashboard.")
        input_username = st.text_input("Username")
        input_password = st.text_input("Password", type="password")
        if st.button("Authenticate", use_container_width=True):
            if input_username in USER_CREDENTIALS and USER_CREDENTIALS[input_username] == input_password:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("❌ Invalid Username or Password.")
    st.stop()

# -----------------------------------------------------------------------------
# 2. Master Data Ingestion & Cross-Referencing Engine
# -----------------------------------------------------------------------------
@st.cache_data
def load_master_data():
    today = pd.to_datetime('today')
    
    # --- A. Load Master Workforce Base ---
    try:
        df_emp = pd.read_excel("Mines_Emp_list.xlsx")
        df_emp.columns = df_emp.columns.astype(str).str.strip()
        df_base = df_emp[['pers_no', 'name', 'desg', 'dept_nm']].copy()
        df_base.rename(columns={'pers_no': 'Pers No', 'name': 'Name', 'desg': 'Designation', 'dept_nm': 'Department'}, inplace=True)
        df_base['Pers No'] = pd.to_numeric(df_base['Pers No'], errors='coerce')
    except Exception as e:
        st.error(f"Failed to load Base Employee List: {e}")
        df_base = pd.DataFrame(columns=['Pers No', 'Name', 'Designation', 'Department'])

    # --- B. Load PME Separately for 2025, 2026, and Latest (Bulletproof Scanner) ---
    def process_pme_file(filename):
        dfs = []
        try:
            xls_pme = pd.ExcelFile(filename)
            for sheet in xls_pme.sheet_names:
                df = pd.read_excel(xls_pme, sheet_name=sheet)
                header_idx = None
                for i, row in df.iterrows():
                    row_str = [str(val).lower().replace(' ', '').replace('.', '') for val in row.dropna().values]
                    if any('plno' in val or 'persno' in val for val in row_str):
                        header_idx = i; break
                        
                if header_idx is not None:
                    df = pd.read_excel(xls_pme, sheet_name=sheet, header=header_idx + 1)
                
                df.columns = df.columns.astype(str).str.strip()
                
                # Auto-correct column name variations
                for col in df.columns:
                    col_clean = col.lower().replace(' ', '').replace('.', '')
                    if 'dateoftest' in col_clean:
                        df.rename(columns={col: 'Date of test'}, inplace=True)
                    elif 'plno' in col_clean or 'persno' in col_clean:
                        df.rename(columns={col: 'Pl.No.'}, inplace=True)
                        
                if 'Pl.No.' in df.columns:
                    dfs.append(df)
        except Exception: pass
        
        if dfs:
            df_raw = pd.concat(dfs, ignore_index=True)
            if 'Pl.No.' in df_raw.columns:
                df_raw['Pl.No.'] = pd.to_numeric(df_raw['Pl.No.'], errors='coerce')
                df_raw = df_raw.dropna(subset=['Pl.No.'])
                if 'Date of test' in df_raw.columns:
                    df_raw['Date of test'] = pd.to_datetime(df_raw['Date of test'], format='%d.%m.%Y', errors='coerce')
                    df_raw = df_raw.sort_values('Date of test').groupby('Pl.No.', as_index=False).last()
                return df_raw
        return pd.DataFrame()

    raw_2025 = process_pme_file("PME 2025.xlsx")
    raw_2026 = process_pme_file("PME 2026.xlsx")
    
    if not raw_2025.empty and not raw_2026.empty:
        raw_latest = pd.concat([raw_2025, raw_2026], ignore_index=True)
    elif not raw_2025.empty:
        raw_latest = raw_2025.copy()
    elif not raw_2026.empty:
        raw_latest = raw_2026.copy()
    else:
        raw_latest = pd.DataFrame()
        
    if not raw_latest.empty and 'Date of test' in raw_latest.columns:
        raw_latest = raw_latest.sort_values('Date of test').groupby('Pl.No.', as_index=False).last()

    def build_pme_module(raw_df):
        if not raw_df.empty and 'Pl.No.' in raw_df.columns and 'Date of test' in raw_df.columns:
            df_out = df_base.merge(raw_df[['Pl.No.', 'Date of test']], left_on='Pers No', right_on='Pl.No.', how='left')
        else:
            df_out = df_base.copy()
            df_out['Date of test'] = pd.NaT
        df_out['Due Date'] = df_out['Date of test'] + pd.DateOffset(years=1)
        df_out['Status'] = df_out['Due Date'].apply(lambda x: 'Valid' if pd.notna(x) and x >= today else 'Overdue')
        return df_out

    df_pme_latest = build_pme_module(raw_latest)
    df_pme_2025 = build_pme_module(raw_2025)
    df_pme_2026 = build_pme_module(raw_2026)

    # --- C. Load Training Desk Modules ---
    master_data = {'Refresher': pd.DataFrame(), 'First_Aid': pd.DataFrame(), 'Supervisor': pd.DataFrame(), 'Fire_Fighting': pd.DataFrame(), 'OEM_Training': pd.DataFrame(), 'Service_Training': pd.DataFrame()}
    try:
        xls = pd.ExcelFile("Training_Desk_Final_OEM_Service_v2 (2).xlsx")
        for sheet in master_data.keys():
            if sheet in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet)
                h_idx = None
                for i, row in df.iterrows():
                    if any(str(val).strip() in ['Pers No', 'Pl.No.', 'Name'] for val in row.dropna().values):
                        h_idx = i; break
                if h_idx is not None:
                    df = pd.read_excel(xls, sheet_name=sheet, header=h_idx + 1)
                
                df.columns = df.columns.astype(str).str.strip()
                if 'Pers No' in df.columns:
                    df['Pers No'] = pd.to_numeric(df['Pers No'], errors='coerce')
                master_data[sheet] = df
    except Exception: pass

    def merge_to_base(sheet_name, cols_to_keep, status_col='Due Alert'):
        df_raw = master_data[sheet_name]
        if not df_raw.empty and 'Pers No' in df_raw.columns:
            avail_cols = [c for c in cols_to_keep if c in df_raw.columns] + ['Pers No']
            df_merged = df_base.merge(df_raw[avail_cols], on='Pers No', how='left')
            if status_col in df_merged.columns:
                df_merged['Status'] = df_merged[status_col].apply(lambda x: 'Overdue' if str(x).strip().lower() == 'due' or pd.isna(x) else 'Valid')
            return df_merged
        else:
            df_empty = df_base.copy()
            df_empty['Status'] = 'Overdue'
            return df_empty

    df_refresher = merge_to_base('Refresher', ['Refresher Last Date', 'Refresher Expiry Date', 'Due Alert'])
    df_firstaid = merge_to_base('First_Aid', ['First Aid Last Year', 'First Aid Type', 'First Aid Expiry Year', 'Due Alert'])
    df_supervisor = merge_to_base('Supervisor', ['Supervisor Last Date', 'Supervisor Expiry Date', 'Due Alert'])
    df_fire = merge_to_base('Fire_Fighting', ['Fire Fighting Last Date', 'Fire Fighting Expiry Date', 'Due Alert'])
    
    # --- DATE CLEANER: Remove 00:00:00 by formatting dates to DD-MM-YYYY ---
    for df in [df_pme_latest, df_pme_2025, df_pme_2026, df_refresher, df_firstaid, df_supervisor, df_fire, master_data['OEM_Training'], master_data['Service_Training']]:
        if not df.empty:
            for col in df.columns:
                if 'date' in col.lower() or pd.api.types.is_datetime64_any_dtype(df[col]):
                    df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%d-%m-%Y').fillna('')

    return df_pme_latest, df_pme_2025, df_pme_2026, df_refresher, df_firstaid, df_supervisor, df_fire, master_data['OEM_Training'], master_data['Service_Training']

df_pme_latest, df_pme_2025, df_pme_2026, df_refresher, df_firstaid, df_supervisor, df_fire, df_oem, df_service = load_master_data()

# -----------------------------------------------------------------------------
# 3. Sidebar Search & Global Navigation
# -----------------------------------------------------------------------------
st.sidebar.title("👤 MineSync Portal")
st.sidebar.markdown("### 🔍 Quick Employee Lookup")
search_query = st.sidebar.text_input("Enter Name or Pers No:", placeholder="e.g. 10395 or Saunta")

if search_query:
    st.sidebar.markdown("#### 📄 Scan Results")
    query_str = str(search_query).lower()
    
    def search_df(df, name):
        if not df.empty and 'Pers No' in df.columns and 'Name' in df.columns:
            matches = df[(df['Pers No'].astype(str).str.lower().str.contains(query_str)) | 
                         (df['Name'].astype(str).str.lower().str.contains(query_str))]
            if not matches.empty:
                for _, row in matches.iterrows():
                    status = row.get('Status', 'Unknown')
                    icon = "🚨" if status == 'Overdue' else "✅"
                    st.sidebar.markdown(f"**{name}:** {icon} {status}")
    
    search_df(df_pme
