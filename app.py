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
        .main { background-color: #f4f6f9; }
        h1, h2, h3 { color: #1e3d59; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
        
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
        
        div[data-testid="stSidebar"] .stButton>button { 
            background: #ff4b4b; height: auto; margin-top: 10px;
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
    
    # --- A. Load Master Workforce Base (The 476 Employees) ---
    try:
        df_emp = pd.read_excel("Mines_Emp_list.xlsx")
        df_emp.columns = df_emp.columns.astype(str).str.strip()
        df_base = df_emp[['pers_no', 'name', 'desg', 'dept_nm']].copy()
        df_base.rename(columns={'pers_no': 'Pers No', 'name': 'Name', 'desg': 'Designation', 'dept_nm': 'Department'}, inplace=True)
        df_base['Pers No'] = pd.to_numeric(df_base['Pers No'], errors='coerce')
    except Exception as e:
        st.error(f"Failed to load Base Employee List: {e}")
        df_base = pd.DataFrame(columns=['Pers No', 'Name', 'Designation', 'Department'])

    # --- B. Load & Merge PME ---
    pme_dfs = []
    for file in ["PME 2025.xlsx", "PME 2026.xlsx"]:
        try:
            xls_pme = pd.ExcelFile(file)
            for sheet in xls_pme.sheet_names:
                df = pd.read_excel(xls_pme, sheet_name=sheet)
                header_idx = None
                for i, row in df.iterrows():
                    if any(str(val).strip() == 'Pl.No.' for val in row.dropna().values):
                        header_idx = i; break
                if header_idx is not None:
                    df = pd.read_excel(xls_pme, sheet_name=sheet, header=header_idx + 1)
                df.columns = df.columns.astype(str).str.strip()
                for col in df.columns:
                    if 'date of test' in col.lower():
                        df.rename(columns={col: 'Date of test'}, inplace=True)
                pme_dfs.append(df)
        except Exception: pass
        
    if pme_dfs:
        df_pme_raw = pd.concat(pme_dfs, ignore_index=True)
        df_pme_raw['Pl.No.'] = pd.to_numeric(df_pme_raw['Pl.No.'], errors='coerce')
        df_pme_raw = df_pme_raw.dropna(subset=['Pl.No.'])
        df_pme_raw['Date of test'] = pd.to_datetime(df_pme_raw['Date of test'], format='%d.%m.%Y', errors='coerce')
        df_pme_raw = df_pme_raw.sort_values('Date of test').groupby('Pl.No.', as_index=False).last()
        
        # Cross-reference with the 476 Master List
        df_pme = df_base.merge(df_pme_raw[['Pl.No.', 'Date of test']], left_on='Pers No', right_on='Pl.No.', how='left')
    else:
        df_pme = df_base.copy()
        df_pme['Date of test'] = pd.NaT

    df_pme['Due Date'] = df_pme['Date of test'] + pd.DateOffset(years=1)
    df_pme['Compliance Year'] = df_pme['Date of test'].dt.year
    df_pme['Compliance Month'] = df_pme['Date of test'].dt.strftime('%b %Y')
    df_pme['Status'] = df_pme['Due Date'].apply(lambda x: 'Valid' if pd.notna(x) and x >= today else 'Overdue')
    
    # --- C. Load & Merge Training Desk Modules ---
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

    # Function to cross-reference other modules against the 476 Master List
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
    
    return df_pme, df_refresher, df_firstaid, df_supervisor, df_fire, master_data['OEM_Training'], master_data['Service_Training']

df_pme, df_refresher, df_firstaid, df_supervisor, df_fire, df_oem, df_service = load_master_data()

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
    
    search_df(df_pme, "PME")
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
        for col in available_cols:
            val = row[col]
            if pd.api.types.is_datetime64_any_dtype(type(val)) and pd.notna(val): val = val.strftime('%d-%m-%Y')
            record_parts.append(f"{val}")
        pdf.cell(200, 8, txt=" | ".join(record_parts), ln=True)
    return pdf.output(dest='S').encode('latin-1')

if page == "Enterprise Overview":
    total_emp = len(df_pme)
    ovd_pme = len(df_pme[df_pme['Status'] == 'Overdue'])
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
    if not df_pme.empty:
        with st.expander("📂 View Merged PME Master Data (Against 476 Base)"): st.dataframe(df_pme)
        st.subheader("Action Required: PME Overdue")
        overdue_pme = df_pme[df_pme['Status'] == 'Overdue']
        if not overdue_pme.empty:
            st.dataframe(overdue_pme.style.apply(highlight_overdue, axis=1))
            pdf_bytes = create_pdf_report(overdue_pme, "PME Overdue Report")
            st.download_button("📄 Download PDF", data=pdf_bytes, file_name="Overdue_PME.pdf", mime="application/pdf")
        else: st.success("All personnel are currently compliant.")

elif page == "Refresher":
    st.title("📚 Statutory Refresher")
    if not df_refresher.empty:
        with st.expander("📂 View Master Data (Against 476 Base)"): st.dataframe(df_refresher)
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
        with st.expander("📂 View Master Data (Against 476 Base)"): st.dataframe(df_firstaid)
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
        if not df_pme.empty and 'Compliance Year' in df_pme.columns:
            st.markdown("#### 🩺 Periodic Medical Examination (PME) Trends")
            col_rep1, col_rep2 = st.columns(2)
            
            with col_rep1:
                valid_years = df_pme.dropna(subset=['Compliance Year'])
                year_summary = valid_years.groupby('Compliance Year').size().reset_index(name='Trained Personnel')
                fig_year = px.bar(year_summary, x='Compliance Year', y='Trained Personnel', 
                                  text_auto=True, color_continuous_scale='Blues', 
                                  title="Year-Wise Compliance Volume")
                fig_year.update_layout(plot_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig_year, use_container_width=True)
                
            with col_rep2:
                valid_months = df_pme.dropna(subset=['Compliance Month'])
                month_summary = valid_months.groupby('Compliance Month').size().reset_index(name='Completed Trainings')
                month_summary['Sort Date'] = pd.to_datetime(month_summary['Compliance Month'], format='%b %Y')
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
