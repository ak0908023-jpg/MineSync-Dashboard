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
