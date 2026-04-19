import streamlit as st

from core.database import init_state, seed_default_departments, set_current_page, sync_page_from_query
from pages.capture import render_capture
from pages.dashboard import render_dashboard
from pages.history import render_history
from pages.tracker import render_tracker
from ui.sidebar import render_sidebar
from ui.styles import inject_styles
from utils.helpers import build_action_dataframe, build_meeting_dataframe


def main():
    st.set_page_config(page_title="MeetIQ", layout="wide")
    inject_styles()
    sync_page_from_query()
    init_state()
    seed_default_departments()

    meetings = st.session_state.meetings
    meeting_df = build_meeting_dataframe(meetings)
    action_df = build_action_dataframe(meetings)

    render_sidebar(st.session_state.current_page, set_current_page)

    if st.session_state.current_page == "Dashboard":
        render_dashboard(meetings, meeting_df, action_df)
    elif st.session_state.current_page == "Tracker":
        render_tracker(meetings)
    elif st.session_state.current_page == "Capture":
        render_capture(meetings, st.session_state.departments)
    elif st.session_state.current_page == "History":
        render_history()


if __name__ == "__main__":
    main()
