import streamlit as st


def render_sidebar(current_page: str, set_current_page) -> None:
    with st.sidebar:
        st.markdown(
            """
            <div class="sidebar-title">AI-Powered Meeting Insight Generator &amp; Action Tracker</div>
            <div class="sidebar-subtitle">for Talentcorp by Zaf &lt;3</div>
            """,
            unsafe_allow_html=True,
        )
        st.button("Dashboard", key="nav_dashboard", use_container_width=True, type="primary" if current_page == "Dashboard" else "secondary", on_click=set_current_page, args=("Dashboard",))
        st.button("Tracker", key="nav_tracker", use_container_width=True, type="primary" if current_page == "Tracker" else "secondary", on_click=set_current_page, args=("Tracker",))
        st.button("Capture", key="nav_capture", use_container_width=True, type="primary" if current_page == "Capture" else "secondary", on_click=set_current_page, args=("Capture",))
        st.button("History", key="nav_history", use_container_width=True, type="primary" if current_page == "History" else "secondary", on_click=set_current_page, args=("History",))
