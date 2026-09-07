import streamlit as st
import pandas as pd
import datetime
import io
import base64
from fpdf import FPDF
import plotly.express as px

# -----------------------------------------------------------------------------
# 1. Dashboard Configuration & Custom Styling
# -----------------------------------------------------------------------------
st.set_page_config(page_title="MineSync Compliance Manager", page_icon="⚙️", layout="wide")

# NALCO Background Image Script
def set_bg(main_bg):
    try:
        with open(main_bg, "rb") as image_file:
            encoded_string = base64.b64encode(image_file.read()).decode()
        st.markdown(
            f"""
            <style>
            .stApp {{
                background-image: url(data:image/jpg;base64,{encoded_string});
                background-size: cover;
                background-position: center;
                background-attachment: fixed;
            }}
            /* Semi-transparent overlay so the dashboard remains readable */
            .main .block-container {{
                background-color: rgba(244, 246, 249, 0.92);
                border-radius: 12px;
                padding: 2rem;
                margin-top: 2rem;
            }}
            </style>
            """,
            unsafe_allow_html=True
        )
    except FileNotFoundError:
        pass # App will load normally if the image is missing

set_bg("nalco_bg.jpg")

st.markdown("""
    <style>
        /* Main background and fonts */
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
    
    search_df(df_pme_latest, "PME")
    search_df(df_refresher, "Refresher")
    search_df(df_firstaid, "First Aid")
    search_df(df_fire, "Fire Fighting")
    st.sidebar.divider()

st.sidebar.button("🔓 Secure Logout", on_click=lambda: st.session_state.update(authenticated=False))

# -----------------------------------------------------------------------------
# 4. Top Grid Navigation System
# -----------------------------------------------------------------------------
if "current_page" not in st.session_state: st.session_state.current_page = "Enterprise Overview"
def nav(page_name): st.session_state.current_page = page_name

st.title("⚙️ Executive Training Dashboard")

nav1, nav2, nav3, nav4 = st.columns(4)
with nav1: st.button("📊 Overview", on_click=nav, args=("Enterprise Overview",), use_container_width=True)
with nav2: st.button("🩺 PME Tracking", on_click=nav, args=("PME",), use_container_width=True)
with nav3: st.button("📚 Refresher", on_click=nav, args=("Refresher",), use_container_width=True)
with nav4: st.button("🚑 First Aid", on_click=nav, args=("First Aid",), use_container_width=True)

nav5, nav6, nav7, nav8 = st.columns(4)
with nav5: st.button("🚜 HEMM", on_click=nav, args=("HEMM",), use_container_width=True)
with nav6: st.button("📋 Supervisor", on_click=nav, args=("Supervisor",), use_container_width=True)
with nav7: st.button("🔥 Fire / ⚙️ OEM", on_click=nav, args=("Misc",), use_container_width=True)
with nav8: st.button("📈 Analytics & AI", on_click=nav, args=("Analytics",), use_container_width=True)

st.divider()

# -----------------------------------------------------------------------------
# 5. Page Modules
# -----------------------------------------------------------------------------
page = st.session_state.current_page

def highlight_overdue(row):
    if row.get('Status') == 'Overdue': return ['background-color: #ffe6e6; color: #b30000; font-weight: bold'] * len(row)
    return [''] * len(row)

def create_pdf_report(dataframe, title):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(200, 10, txt=title, ln=True, align='C')
    pdf.set_font("Arial", size=10)
    pdf.cell(200, 10, txt=f"Generated on: {datetime.date.today().strftime('%d-%m-%Y')}", ln=True, align='C')
    pdf.ln(5)
    
    available_cols = [c for c in ['Pers No', 'Name', 'Due Date', 'Refresher Expiry Date', 'First Aid Expiry Year'] if c in dataframe.columns][:3]
    for i, row in dataframe.iterrows():
        record_parts = []
        for col in available_cols: record_parts.append(f"{row[col]}")
        pdf.cell(200, 8, txt=" | ".join(record_parts), ln=True)
    return pdf.output(dest='S').encode('latin-1')

if page == "Enterprise Overview":
    total_emp = len(df_pme_latest)
    ovd_pme = len(df_pme_latest[df_pme_latest['Status'] == 'Overdue'])
    ovd_ref = len(df_refresher[df_refresher['Status'] == 'Overdue'])
    ovd_fa = len(df_firstaid[df_firstaid['Status'] == 'Overdue'])
    ovd_fire = len(df_fire[df_fire['Status'] == 'Overdue'])

    st.markdown("### 🌐 Operations Command Center")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Master Workforce Base", f"{total_emp:,}")
    col2.metric("PME Alerts 🚨", f"{ovd_pme}", delta="Action Required", delta_color="inverse")
    col3.metric("Refresher Alerts 🚨", f"{ovd_ref}", delta="Action Required", delta_color="inverse")
    col4.metric("First Aid Alerts 🚨", f"{ovd_fa}", delta="Action Required", delta_color="inverse")
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    chart_col1, chart_col2 = st.columns([2, 1])
    with chart_col1:
        alert_data = pd.DataFrame({
            'Category': ['PME', 'Refresher', 'First Aid', 'Fire Fighting'],
            'Overdue Count': [ovd_pme, ovd_ref, ovd_fa, ovd_fire]
        })
        fig_alerts = px.bar(alert_data, x='Category', y='Overdue Count', text='Overdue Count',
                            title="Enterprise Vulnerability Map (Overdue Trainings)",
                            color='Overdue Count', color_continuous_scale='Reds')
        fig_alerts.update_traces(textposition='outside', marker_line_color='black', marker_line_width=1.5)
        fig_alerts.update_layout(plot_bgcolor="rgba(0,0,0,0)", margin=dict(t=40, b=20, l=0, r=0))
        st.plotly_chart(fig_alerts, use_container_width=True)

    with chart_col2:
        if total_emp > 0:
            fig_health = px.pie(names=['Compliant', 'Overdue'], values=[total_emp - ovd_pme, ovd_pme], 
                                hole=0.7, title="Global PME Health Index",
                                color_discrete_sequence=['#2ca02c', '#ff4b4b'])
            fig_health.update_layout(margin=dict(t=40, b=20, l=0, r=0))
            st.plotly_chart(fig_health, use_container_width=True)

elif page == "PME":
    st.title("🩺 Periodic Medical Examination")
    tab_latest, tab_2026, tab_2025 = st.tabs(["🌟 Current Status", "📅 2026 Records", "📅 2025 Records"])
    
    with tab_latest:
        st.markdown("#### Master PME Tracking (Active Deadlines)")
        if not df_pme_latest.empty:
            with st.expander("📂 View Merged Latest Data (Against Base)"): st.dataframe(df_pme_latest)
            st.subheader("Action Required: PME Overdue")
            overdue_pme = df_pme_latest[df_pme_latest['Status'] == 'Overdue']
            if not overdue_pme.empty:
                st.dataframe(overdue_pme.style.apply(highlight_overdue, axis=1))
                pdf_bytes = create_pdf_report(overdue_pme, "PME Overdue Report")
                st.download_button("📄 Download PDF", data=pdf_bytes, file_name="Overdue_PME.pdf", mime="application/pdf")
            else: st.success("All personnel are currently compliant.")
            
    with tab_2026:
        st.markdown("#### 2026 PME Completion Overview")
        if not df_pme_2026.empty:
            total_26 = len(df_pme_2026)
            tested_26 = sum(df_pme_2026['Date of test'] != '')
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Base Roster", total_26)
            c2.metric("Tested in 2026", tested_26)
            c3.metric("Missing 2026 Test", total_26 - tested_26)
            c4.metric("2026 Compliance", f"{(tested_26 / total_26 * 100):.1f}%" if total_26 > 0 else "0%")
            st.dataframe(df_pme_2026)
            
    with tab_2025:
        st.markdown("#### 2025 PME Historical Overview")
        if not df_pme_2025.empty:
            total_25 = len(df_pme_2025)
            tested_25 = sum(df_pme_2025['Date of test'] != '')
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Base Roster", total_25)
            c2.metric("Tested in 2025", tested_25)
            c3.metric("Did Not Test", total_25 - tested_25)
            c4.metric("2025 Compliance Rate", f"{(tested_25 / total_25 * 100):.1f}%" if total_25 > 0 else "0%")
            st.dataframe(df_pme_2025)

elif page == "Refresher":
    st.title("📚 Statutory Refresher")
    if not df_refresher.empty:
        with st.expander("📂 View Master Data (Against Base)"): st.dataframe(df_refresher)
        st.subheader("Action Required: Overdue")
        overdue = df_refresher[df_refresher['Status'] == 'Overdue']
        if not overdue.empty: 
            st.dataframe(overdue.style.apply(highlight_overdue, axis=1))
            pdf_bytes = create_pdf_report(overdue, "Refresher Overdue Report")
            st.download_button("📄 Download PDF", data=pdf_bytes, file_name="Overdue_Refresher.pdf", mime="application/pdf")
        else: st.success("All personnel compliant.")

elif page == "First Aid":
    st.title("🚑 First Aid Tracker")
    if not df_firstaid.empty:
        with st.expander("📂 View Master Data (Against Base)"): st.dataframe(df_firstaid)
        st.subheader("Action Required: Overdue")
        overdue = df_firstaid[df_firstaid['Status'] == 'Overdue']
        if not overdue.empty: 
            st.dataframe(overdue.style.apply(highlight_overdue, axis=1))
            pdf_bytes = create_pdf_report(overdue, "First Aid Overdue Report")
            st.download_button("📄 Download PDF", data=pdf_bytes, file_name="Overdue_First_Aid.pdf", mime="application/pdf")

elif page == "HEMM":
    st.title("🚜 HEMM Competency")
    st.info("No HEMM data tab found in the master file.")

elif page == "Supervisor":
    st.title("📋 Supervisor Certification")
    if not df_supervisor.empty: 
        with st.expander("📂 View Data"): st.dataframe(df_supervisor)
        overdue = df_supervisor[df_supervisor['Status'] == 'Overdue']
        if not overdue.empty: st.dataframe(overdue.style.apply(highlight_overdue, axis=1))

elif page == "Misc":
    tab_fire, tab_oem, tab_service = st.tabs(["🔥 Fire Fighting", "⚙️ OEM", "🔧 Service"])
    with tab_fire: 
        if not df_fire.empty: st.dataframe(df_fire)
    with tab_oem: 
        if not df_oem.empty: st.dataframe(df_oem)
    with tab_service: 
        if not df_service.empty: st.dataframe(df_service)

elif page == "Analytics":
    tab_charts, tab_ai = st.tabs(["📈 Data Analytics", "💬 AI Assistant"])
    
    with tab_charts:
        df_chart = df_pme_latest.copy()
        df_chart['Compliance Year'] = pd.to_datetime(df_chart['Date of test'], format='%d-%m-%Y', errors='coerce').dt.year
        df_chart['Compliance Month'] = pd.to_datetime(df_chart['Date of test'], format='%d-%m-%Y', errors='coerce').dt.strftime('%b %Y')
        
        if not df_chart.empty and 'Compliance Year' in df_chart.columns:
            st.markdown("#### 🩺 Periodic Medical Examination (PME) Trends")
            col_rep1, col_rep2 = st.columns(2)
            
            with col_rep1:
                valid_years = df_chart.dropna(subset=['Compliance Year'])
                year_summary = valid_years.groupby('Compliance Year').size().reset_index(name='Trained Personnel')
                fig_year = px.bar(year_summary, x='Compliance Year', y='Trained Personnel', 
                                  text_auto=True, color_continuous_scale='Blues', 
                                  title="Year-Wise Compliance Volume")
                fig_year.update_layout(plot_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig_year, use_container_width=True)
                
            with col_rep2:
                valid_months = df_chart.dropna(subset=['Compliance Month'])
                month_summary = valid_months.groupby('Compliance Month').size().reset_index(name='Completed Trainings')
                month_summary['Sort Date'] = pd.to_datetime(month_summary['Compliance Month'], format='%b %Y', errors='coerce')
                month_summary = month_summary.sort_values('Sort Date')
                fig_month = px.line(month_summary, x='Compliance Month', y='Completed Trainings', 
                                    markers=True, title="Month-Wise Due vs. Compliance", 
                                    color_discrete_sequence=['#ff4b4b'])
                fig_month.update_layout(plot_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig_month, use_container_width=True)

        st.divider()

        if not df_firstaid.empty and 'First Aid Last Year' in df_firstaid.columns:
            st.markdown("#### 🚑 First Aid Training Breakdowns")
            col_fa1, col_fa2 = st.columns(2)
            
            with col_fa1:
                valid_fa_years = df_firstaid.dropna(subset=['First Aid Last Year'])
                fa_year_summary = valid_fa_years.groupby('First Aid Last Year').size().reset_index(name='Trained Personnel')
                fig_fa_bar = px.bar(fa_year_summary, x='First Aid Last Year', y='Trained Personnel', 
                                    text_auto=True, title="Annual First Aid Certifications", 
                                    color_discrete_sequence=['#2ca02c'])
                fig_fa_bar.update_layout(plot_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig_fa_bar, use_container_width=True)
                
            with col_fa2:
                if 'First Aid Type' in df_firstaid.columns:
                    fa_type_summary = df_firstaid.groupby('First Aid Type').size().reset_index(name='Total Trained')
                    fig_fa_pie = px.pie(fa_type_summary, values='Total Trained', names='First Aid Type', 
                                        hole=0.4, title="Breakdown by First Aid Type", 
                                        color_discrete_sequence=px.colors.qualitative.Set2)
                    st.plotly_chart(fig_fa_pie, use_container_width=True)

    with tab_ai:
        st.markdown("Ask questions about training schedules, statutory compliance, or dashboard data.")
        if "messages" not in st.session_state: 
            st.session_state.messages = [{"role": "assistant", "content": "How can I help you today?"}]
        for message in st.session_state.messages:
            with st.chat_message(message["role"]): 
                st.markdown(message["content"])
        if prompt := st.chat_input("E.g., How many people are overdue for PME?"):
            with st.chat_message("user"): 
                st.markdown(prompt)
            st.session_state.messages.append({"role": "user", "content": prompt})
            response = f"You asked: '{prompt}'. I am currently in demonstration mode."
            with st.chat_message("assistant"): 
                st.markdown(response)
            st.session_state.messages.append({"role": "assistant", "content": response})
