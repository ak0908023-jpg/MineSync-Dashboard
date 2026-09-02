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
        .main { background-color: #f8f9fa; }
        h1, h2, h3 { color: #1e3d59; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
        
        div[data-testid="metric-container"] {
            background-color: #ffffff; border-left: 5px solid #1e3d59;
            padding: 15px 20px; border-radius: 8px; box-shadow: 2px 2px 10px rgba(0,0,0,0.05);
        }
        
        .stButton>button {
            background-color: #1e3d59; color: white; border-radius: 8px; border: none;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1); transition: all 0.3s ease;
            height: 60px; font-size: 16px; font-weight: bold;
        }
        .stButton>button:hover { background-color: #3b6998; transform: translateY(-2px); }
        div[data-testid="stSidebar"] .stButton>button { background-color: #ff4b4b; height: auto; }
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
    st.markdown("<h2 style='text-align: center; color: #1e3d59;'>🔒 MineSync Secure Login</h2>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.info("Please enter your credentials to access the compliance dashboard.")
        input_username = st.text_input("Username")
        input_password = st.text_input("Password", type="password")
        if st.button("Secure Login", use_container_width=True):
            if input_username in USER_CREDENTIALS and USER_CREDENTIALS[input_username] == input_password:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("❌ Invalid Username or Password. Please try again.")
    st.stop()

st.sidebar.title("👤 Amit")
st.sidebar.button("🔓 Secure Logout", on_click=lambda: st.session_state.update(authenticated=False))
st.sidebar.divider()

FILE_PATH = "Master_Training_Report_2026.xlsx"

# -----------------------------------------------------------------------------
# 2. Master Data Ingestion & Optimization (Cached)
# -----------------------------------------------------------------------------
@st.cache_data
def load_master_data():
    try:
        xls = pd.ExcelFile(FILE_PATH)
        today = pd.to_datetime('today')
        
        df_pme = pd.read_excel(xls, sheet_name="PME", header=1) if "PME" in xls.sheet_names else pd.DataFrame()
        if not df_pme.empty:
            df_pme.columns = df_pme.columns.astype(str).str.strip()
            if 'Date of test' in df_pme.columns:
                df_pme['Date of test'] = pd.to_datetime(df_pme['Date of test'], format='%d.%m.%Y', errors='coerce')
                df_pme['Due Date'] = df_pme['Date of test'] + pd.DateOffset(years=1)
                df_pme['Compliance Year'] = df_pme['Date of test'].dt.year
                df_pme['Compliance Month'] = df_pme['Date of test'].dt.strftime('%b %Y')
                df_pme['Status'] = df_pme['Due Date'].apply(lambda x: 'Overdue' if pd.notna(x) and x < today else 'Valid')
            else:
                df_pme['Due Date'] = pd.NaT
                df_pme['Status'] = 'Unknown'

        df_refresher = pd.read_excel(xls, sheet_name="Refresher", header=1) if "Refresher" in xls.sheet_names else pd.DataFrame()
        if not df_refresher.empty:
            df_refresher.columns = df_refresher.columns.astype(str).str.strip()
            if 'Date of test' in df_refresher.columns:
                df_refresher['Date of test'] = pd.to_datetime(df_refresher['Date of test'], format='%d.%m.%Y', errors='coerce')
                df_refresher['Due Date'] = df_refresher['Date of test'] + pd.DateOffset(years=4)
                df_refresher['Status'] = df_refresher['Due Date'].apply(lambda x: 'Overdue' if pd.notna(x) and x < today else 'Valid')
            else:
                df_refresher['Due Date'] = pd.NaT
                df_refresher['Status'] = 'Unknown'

        df_firstaid = pd.read_excel(xls, sheet_name="First_Aid", header=1) if "First_Aid" in xls.sheet_names else pd.DataFrame()
        if not df_firstaid.empty:
            df_firstaid.columns = df_firstaid.columns.astype(str).str.strip()
            if 'Date of test' in df_firstaid.columns:
                df_firstaid['Date of test'] = pd.to_datetime(df_firstaid['Date of test'], format='%d.%m.%Y', errors='coerce')
                df_firstaid['Last First Aid Year'] = df_firstaid['Date of test'].dt.year
            else:
                df_firstaid['Last First Aid Year'] = pd.NaT

        df_supervisor = pd.read_excel(xls, sheet_name="Supervisor", header=1) if "Supervisor" in xls.sheet_names else pd.DataFrame()
        df_fire = pd.read_excel(xls, sheet_name="Fire_Fighting") if "Fire_Fighting" in xls.sheet_names else pd.DataFrame()
        df_oem = pd.read_excel(xls, sheet_name="OEM_Training") if "OEM_Training" in xls.sheet_names else pd.DataFrame()
        df_service = pd.read_excel(xls, sheet_name="Service_Training") if "Service_Training" in xls.sheet_names else pd.DataFrame()
        
        return df_pme, df_refresher, df_firstaid, df_supervisor, df_fire, df_oem, df_service
        
    except Exception as e:
        st.error(f"Error loading data: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

df_pme, df_refresher, df_firstaid, df_supervisor, df_fire, df_oem, df_service = load_master_data()

def highlight_overdue(row):
    if row.get('Status') == 'Overdue':
        return ['background-color: #ffcccc; color: #900000; font-weight: bold'] * len(row)
    return [''] * len(row)

def create_pdf_report(dataframe, title):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(200, 10, txt=title, ln=True, align='C')
    pdf.set_font("Arial", size=10)
    pdf.cell(200, 10, txt=f"Generated on: {datetime.date.today().strftime('%d-%m-%Y')}", ln=True, align='C')
    pdf.ln(5)
    
    available_cols = [c for c in ['Pl.No.', 'Name (Shri/Smt)', 'Due Date'] if c in dataframe.columns]
    for i, row in dataframe.iterrows():
        record_parts = []
        for col in available_cols:
            val = row[col]
            if pd.api.types.is_datetime64_any_dtype(type(val)) and pd.notna(val): val = val.strftime('%d-%m-%Y')
            record_parts.append(f"{val}")
        pdf.cell(200, 8, txt=" | ".join(record_parts), ln=True)
    return pdf.output(dest='S').encode('latin-1')

# -----------------------------------------------------------------------------
# 3. Top Grid Navigation System
# -----------------------------------------------------------------------------
if "current_page" not in st.session_state: st.session_state.current_page = "Enterprise Overview"

def nav(page_name): st.session_state.current_page = page_name

st.title("⚙️ DGMS Training Dashboard")

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
# 4. Page Modules (Optimized for Cloud Performance)
# -----------------------------------------------------------------------------
page = st.session_state.current_page

if page == "Enterprise Overview":
    total_personnel = len(df_pme) if not df_pme.empty else 0
    overdue_pme_count = len(df_pme[df_pme['Status'] == 'Overdue']) if not df_pme.empty else 0
    overdue_refresher_count = len(df_refresher[df_refresher['Status'] == 'Overdue']) if not df_refresher.empty else 0

    st.markdown("### 📊 Live Action Requirements")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Workforce", f"{total_personnel:,}")
    col2.metric("PME Overdue 🚨", f"{overdue_pme_count}", delta=f"-{overdue_pme_count} Action Req", delta_color="inverse")
    col3.metric("Refresher Overdue 🚨", f"{overdue_refresher_count}", delta=f"-{overdue_refresher_count} Action Req", delta_color="inverse")
    col4.metric("System Status", "Online ✅")
    
    if total_personnel > 0:
        fig_health = px.pie(names=['Compliant', 'Overdue'], values=[total_personnel - overdue_pme_count, overdue_pme_count], hole=0.6, color_discrete_sequence=['#2ca02c', '#ff4b4b'])
        fig_health.update_layout(margin=dict(t=20, b=20, l=0, r=0), height=300)
        st.plotly_chart(fig_health, use_container_width=True)

elif page == "PME":
    st.title("🩺 Periodic Medical Examination (PME)")
    if not df_pme.empty:
        # Optimization: Only load the massive table if the user clicks the expander
        with st.expander("📂 View Full PME Master Data (May take a moment to load)"):
            st.dataframe(df_pme)
            
        st.subheader("Action Required: PME Overdue")
        overdue_pme = df_pme[df_pme['Status'] == 'Overdue']
        if not overdue_pme.empty:
            st.dataframe(overdue_pme.style.apply(highlight_overdue, axis=1))
            pdf_bytes = create_pdf_report(overdue_pme, "Overdue PME Report")
            st.download_button("📄 Download Overdue PDF", data=pdf_bytes, file_name="Overdue_PME.pdf", mime="application/pdf")
        else:
            st.success("All personnel are currently compliant.")

elif page == "Refresher":
    st.title("📚 4-Year Refresher Scheduling")
    if not df_refresher.empty:
        with st.expander("📂 View Full Refresher Master Data"):
            st.dataframe(df_refresher)
            
        st.subheader("Action Required: Refresher Overdue")
        overdue_refresher = df_refresher[df_refresher['Status'] == 'Overdue']
        if not overdue_refresher.empty:
            st.dataframe(overdue_refresher.style.apply(highlight_overdue, axis=1))
            pdf_bytes = create_pdf_report(overdue_refresher, "Overdue Refresher Report")
            st.download_button("📄 Download Overdue PDF", data=pdf_bytes, file_name="Overdue_Refresher.pdf", mime="application/pdf")
        else:
            st.success("All personnel are currently compliant.")

elif page == "First Aid":
    st.title("🚑 First Aid Training Tracker")
    if not df_firstaid.empty: st.dataframe(df_firstaid)

elif page == "HEMM":
    st.title("🚜 HEMM Operator Competency")
    equipment_filter = st.selectbox("Filter by Equipment Type:", ["All", "Dumper", "Dozer", "Excavator", "Grader"])

elif page == "Supervisor":
    st.title("📋 Supervisor Training & Certification")
    if not df_supervisor.empty: st.dataframe(df_supervisor)

elif page == "Misc":
    tab_fire, tab_oem, tab_service = st.tabs(["🔥 Fire Fighting", "⚙️ OEM Training", "🔧 Service & Maintenance"])
    with tab_fire: 
        if not df_fire.empty: st.dataframe(df_fire)
    with tab_oem: 
        if not df_oem.empty: st.dataframe(df_oem)
    with tab_service: 
        if not df_service.empty: st.dataframe(df_service)

elif page == "Analytics":
    tab_charts, tab_ai = st.tabs(["📈 Data Analytics", "💬 AI Assistant"])
    
    with tab_charts:
        if not df_pme.empty and 'Compliance Year' in df_pme.columns:
            st.markdown("#### 🩺 Periodic Medical Examination (PME) Trends")
            valid_years = df_pme.dropna(subset=['Compliance Year'])
            year_summary = valid_years.groupby('Compliance Year').size().reset_index(name='Trained Personnel')
            fig_year = px.bar(year_summary, x='Compliance Year', y='Trained Personnel', text_auto=True, color_continuous_scale='Blues')
            st.plotly_chart(fig_year, use_container_width=True)

        if not df_firstaid.empty and 'Last First Aid Year' in df_firstaid.columns:
            st.markdown("#### 🚑 First Aid Training Breakdowns")
            valid_fa_years = df_firstaid.dropna(subset=['Last First Aid Year'])
            fa_year_summary = valid_fa_years.groupby('Last First Aid Year').size().reset_index(name='Trained Personnel')
            fig_fa_bar = px.bar(fa_year_summary, x='Last First Aid Year', y='Trained Personnel', text_auto=True, color_discrete_sequence=['#2ca02c'])
            st.plotly_chart(fig_fa_bar, use_container_width=True)

    with tab_ai:
        st.markdown("Ask questions about training schedules, statutory compliance, or dashboard data.")
        if "messages" not in st.session_state: st.session_state.messages = [{"role": "assistant", "content": "How can I help you today?"}]
        for message in st.session_state.messages:
            with st.chat_message(message["role"]): st.markdown(message["content"])
        if prompt := st.chat_input("E.g., How many people are overdue for PME?"):
            with st.chat_message("user"): st.markdown(prompt)
            st.session_state.messages.append({"role": "user", "content": prompt})
            response = f"You asked: '{prompt}'. I am currently in demonstration mode."
            with st.chat_message("assistant"): st.markdown(response)
            st.session_state.messages.append({"role": "assistant", "content": response})
