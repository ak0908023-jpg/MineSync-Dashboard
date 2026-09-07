import streamlit as st
import pandas as pd
import datetime
import io
from fpdf import FPDF
import plotly.express as px
import plotly.graph_objects as go

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
            border-top: 1px solid #e0e0e0;
            border-right: 1px solid #e0e0e0;
            border-bottom: 1px solid #e0e0e0;
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
            letter-spacing: 0.5px;
        }
        .stButton>button:hover { 
            background: linear-gradient(to right, #2b577d, #3b6998); 
            transform: translateY(-2px); 
            box-shadow: 0 6px 15px rgba(0,0,0,0.2);
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
# 2. Master Data Ingestion & Optimization (Cached)
# -----------------------------------------------------------------------------
MASTER_FILE = "Training_Desk_Final_OEM_Service_v2 (2).xlsx"
PME_FILES = ["PME 2025.xlsx", "PME 2026.xlsx"]

@st.cache_data
def load_master_data():
    today = pd.to_datetime('today')
    pme_dfs = []
    
    # 1. Load PME
    for file in PME_FILES:
        try:
            xls_pme = pd.ExcelFile(file)
            for sheet in xls_pme.sheet_names:
                df = pd.read_excel(xls_pme, sheet_name=sheet)
                header_idx = None
                for i, row in df.iterrows():
                    if any(str(val).strip() == 'Pl.No.' for val in row.dropna().values):
                        header_idx = i
                        break
                if header_idx is not None:
                    df = pd.read_excel(xls_pme, sheet_name=sheet, header=header_idx + 1)
                
                df.columns = df.columns.astype(str).str.strip()
                pme_dfs.append(df)
        except FileNotFoundError:
            pass

    if pme_dfs:
        df_pme = pd.concat(pme_dfs, ignore_index=True)
        df_pme = df_pme.dropna(subset=['Pl.No.'])
        if 'Date of test' in df_pme.columns:
            df_pme['Date of test'] = pd.to_datetime(df_pme['Date of test'], format='%d.%m.%Y', errors='coerce')
            df_pme = df_pme.sort_values('Date of test').groupby('Pl.No.', as_index=False).last()
            df_pme['Due Date'] = df_pme['Date of test'] + pd.DateOffset(years=1)
            df_pme['Compliance Year'] = df_pme['Date of test'].dt.year
            df_pme['Compliance Month'] = df_pme['Date of test'].dt.strftime('%b %Y')
            df_pme['Status'] = df_pme['Due Date'].apply(lambda x: 'Overdue' if pd.notna(x) and x < today else 'Valid')
    else:
        df_pme = pd.DataFrame()

    # 2. Load Master Training Desk File
    master_data = {
        'Refresher': pd.DataFrame(), 'First_Aid': pd.DataFrame(), 
        'Supervisor': pd.DataFrame(), 'Fire_Fighting': pd.DataFrame(), 
        'OEM_Training': pd.DataFrame(), 'Service_Training': pd.DataFrame()
    }
    try:
        xls = pd.ExcelFile(MASTER_FILE)
        for sheet in master_data.keys():
            if sheet in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet)
                h_idx = None
                for i, row in df.iterrows():
                    if any(str(val).strip() in ['Pers No', 'Pl.No.', 'Name'] for val in row.dropna().values):
                        h_idx = i
                        break
                if h_idx is not None:
                    df = pd.read_excel(xls, sheet_name=sheet, header=h_idx + 1)
                
                df.columns = df.columns.astype(str).str.strip()
                if 'Due Alert' in df.columns:
                    df['Status'] = df['Due Alert'].apply(lambda x: 'Overdue' if str(x).strip().lower() == 'due' else 'Valid')
                master_data[sheet] = df
    except Exception as e:
        pass
        
    return df_pme, master_data['Refresher'], master_data['First_Aid'], master_data['Supervisor'], master_data['Fire_Fighting'], master_data['OEM_Training'], master_data['Service_Training']

df_pme, df_refresher, df_firstaid, df_supervisor, df_fire, df_oem, df_service = load_master_data()

# -----------------------------------------------------------------------------
# 3. Sidebar Search & Global Navigation
# -----------------------------------------------------------------------------
st.sidebar.title("👤 MineSync Portal")

# THE NEW INFORMATION / SEARCH BAR
st.sidebar.markdown("### 🔍 Quick Employee Lookup")
search_query = st.sidebar.text_input("Enter Name or Pers No:", placeholder="e.g. 10395 or Saunta")

if search_query:
    st.sidebar.markdown("#### 📄 Scan Results")
    query_str = str(search_query).lower()
    
    def search_df(df, name, p_col, n_col):
        if not df.empty and p_col in df.columns and n_col in df.columns:
            matches = df[(df[p_col].astype(str).str.lower().str.contains(query_str)) | 
                         (df[n_col].astype(str).str.lower().str.contains(query_str))]
            if not matches.empty:
                for _, row in matches.iterrows():
                    status = row.get('Status', 'Unknown')
                    icon = "🚨" if status == 'Overdue' else "✅"
                    st.sidebar.markdown(f"**{name}:** {icon} {status}")
    
    search_df(df_pme, "PME", "Pl.No.", "Name (Shri/Smt)")
    search_df(df_refresher, "Refresher", "Pers No", "Name")
    search_df(df_firstaid, "First Aid", "Pers No", "Name")
    search_df(df_fire, "Fire Fighting", "Pers No", "Name")
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
# 5. Page Modules (Stylish Enhancements)
# -----------------------------------------------------------------------------
page = st.session_state.current_page

def highlight_overdue(row):
    if row.get('Status') == 'Overdue': return ['background-color: #ffe6e6; color: #b30000; font-weight: bold'] * len(row)
    return [''] * len(row)

if page == "Enterprise Overview":
    # Calculate all metrics dynamically
    total_pme = len(df_pme)
    ovd_pme = len(df_pme[df_pme['Status'] == 'Overdue']) if not df_pme.empty else 0
    ovd_ref = len(df_refresher[df_refresher['Status'] == 'Overdue']) if not df_refresher.empty else 0
    ovd_fa = len(df_firstaid[df_firstaid['Status'] == 'Overdue']) if not df_firstaid.empty else 0
    ovd_fire = len(df_fire[df_fire['Status'] == 'Overdue']) if not df_fire.empty else 0

    st.markdown("### 🌐 Operations Command Center")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Workforce", f"{total_pme:,}")
    col2.metric("PME Alerts 🚨", f"{ovd_pme}", delta="Action Required", delta_color="inverse")
    col3.metric("Refresher Alerts 🚨", f"{ovd_ref}", delta="Action Required", delta_color="inverse")
    col4.metric("First Aid Alerts 🚨", f"{ovd_fa}", delta="Action Required", delta_color="inverse")
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    chart_col1, chart_col2 = st.columns([2, 1])
    
    with chart_col1:
        # Consolidated Overdue Bar Chart
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
        if total_pme > 0:
            fig_health = px.pie(names=['Compliant', 'Overdue'], values=[total_pme - ovd_pme, ovd_pme], 
                                hole=0.7, title="Global PME Health Index",
                                color_discrete_sequence=['#2ca02c', '#ff4b4b'])
            fig_health.update_layout(margin=dict(t=40, b=20, l=0, r=0))
            st.plotly_chart(fig_health, use_container_width=True)

# THE REST OF THE PAGE MODULES REMAIN THE SAME...
elif page == "PME":
    st.title("🩺 Periodic Medical Examination")
    if not df_pme.empty:
        with st.expander("📂 View Merged PME Master Data"): st.dataframe(df_pme)
        st.subheader("Action Required: PME Overdue")
        overdue_pme = df_pme[df_pme['Status'] == 'Overdue']
        if not overdue_pme.empty:
            st.dataframe(overdue_pme.style.apply(highlight_overdue, axis=1))
        else: st.success("All personnel are currently compliant.")

elif page == "Refresher":
    st.title("📚 Statutory Refresher")
    if not df_refresher.empty:
        with st.expander("📂 View Master Data"): st.dataframe(df_refresher)
        st.subheader("Action Required: Overdue")
        if 'Status' in df_refresher.columns:
            overdue = df_refresher[df_refresher['Status'] == 'Overdue']
            if not overdue.empty: st.dataframe(overdue.style.apply(highlight_overdue, axis=1))
            else: st.success("All personnel compliant.")

elif page == "First Aid":
    st.title("🚑 First Aid Tracker")
    if not df_firstaid.empty:
        with st.expander("📂 View Master Data"): st.dataframe(df_firstaid)
        if 'Status' in df_firstaid.columns:
            overdue = df_firstaid[df_firstaid['Status'] == 'Overdue']
            if not overdue.empty: st.dataframe(overdue.style.apply(highlight_overdue, axis=1))

elif page == "HEMM":
    st.title("🚜 HEMM Competency")
    st.info("No HEMM data tab found in the master file.")

elif page == "Supervisor":
    st.title("📋 Supervisor Certification")
    if not df_supervisor.empty: 
        with st.expander("📂 View Data"): st.dataframe(df_supervisor)
        if 'Status' in df_supervisor.columns:
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
        st.info("Advanced analytics engine ready.")
    with tab_ai:
        st.markdown("Ask questions about training schedules or dashboard data.")
        if "messages" not in st.session_state: st.session_state.messages = [{"role": "assistant", "content": "How can I help you today?"}]
        for message in st.session_state.messages:
            with st.chat_message(message["role"]): st.markdown(message["content"])
        if prompt := st.chat_input("Ask the Assistant..."):
            with st.chat_message("user"): st.markdown(prompt)
            st.session_state.messages.append({"role": "user", "content": prompt})
            with st.chat_message("assistant"): st.markdown(f"I am in demo mode. You asked: {prompt}")
            st.session_state.messages.append({"role": "assistant", "content": f"I am in demo mode. You asked: {prompt}"})
