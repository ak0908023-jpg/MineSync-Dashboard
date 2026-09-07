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
                    df = pd.read_
