import calendar
import sys
from datetime import date
from pathlib import Path

import streamlit as st

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.database import set_tracker_shortcut
from core.pipeline import chat_with_meetings
from ui.calendar import build_calendar_html
from ui.components import render_action_card, render_chat_bubble, render_completion_ring, render_kpi_card
from utils.formatters import normalize_status, normalize_value
from utils.helpers import get_upcoming_meetings


def render_dashboard_chat(meetings: list) -> None:
    current_user_id = st.session_state.get("chat_user_id", "").strip()
    st.text_input("User ID", key="chat_user_id", placeholder="Enter your ID before chatting", help="Chat history is private to this ID.")
    if not current_user_id:
        st.info("Enter your User ID to unlock private chat and history.")
        return
    if st.session_state.get("active_chat_user_id") != current_user_id:
        st.session_state.active_chat_user_id = current_user_id
        st.session_state.chat_history = []

    chat_history_box = st.container(height=420, border=False)
    with chat_history_box:
        st.markdown('<div class="chat-thread dashboard-chat-thread">', unsafe_allow_html=True)
        if st.session_state.chat_history:
            for message in st.session_state.chat_history:
                render_chat_bubble(message["role"], message["text"])
        else:
            render_chat_bubble("assistant", "Ask about meetings, tasks, deadlines, or next steps.")
        st.markdown("</div>", unsafe_allow_html=True)

    with st.form("dashboard_chat_form", clear_on_submit=True):
        question = st.text_input("Ask AI", placeholder="Ask about a meeting or task...", label_visibility="collapsed")
        submitted = st.form_submit_button("Send", use_container_width=True)

    if submitted and question.strip():
        st.session_state.chat_history.append({"role": "user", "text": question})
        try:
            answer = chat_with_meetings(question, meetings)
        except Exception as exc:
            answer = f"Error: {exc}"
        st.session_state.chat_history.append({"role": "assistant", "text": answer})
        st.rerun()


def render_dashboard(meetings: list, meeting_df, action_df) -> None:
    dashboard_years = sorted(meeting_df["year"].dropna().unique().tolist(), reverse=True) if not meeting_df.empty else [date.today().year]
    all_actions = [action for meeting in meetings for action in meeting.get("actions", [])]
    normalized_action_statuses = [normalize_status(action) for action in all_actions]
    total_action_items = len(all_actions)
    pending_action_items = sum(1 for status in normalized_action_statuses if status in {"Pending", "In Progress", "Overdue"})
    completed_action_items = sum(1 for status in normalized_action_statuses if status == "Done")
    completion_pct = round((completed_action_items / total_action_items) * 100) if total_action_items else 0

    dashboard_left, dashboard_right = st.columns([1.35, 0.85])
    with dashboard_left:
        overview_card = st.container(border=True)
        with overview_card:
            st.markdown("### Today's Brief")
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                render_kpi_card("Total Meeting", str(len(meetings)), "Stored records", "#0f766e")
            with c2:
                render_kpi_card("Total Action Item", str(total_action_items), "Tracked tasks", "#1e3a5f")
            with c3:
                render_kpi_card("Follow up needed", str(pending_action_items), "Open tasks", "#d97706")
            with c4:
                render_completion_ring(completion_pct)

        quick_nav = st.container(border=True)
        with quick_nav:
            st.markdown("### Quick View")
            b1, b2 = st.columns(2)
            with b1:
                if st.button("Open Pending Tracker", use_container_width=True):
                    set_tracker_shortcut("open")
                    st.rerun()
            with b2:
                if st.button("Open Completed Tracker", use_container_width=True):
                    set_tracker_shortcut("done")
                    st.rerun()

        upcoming_card = st.container(border=True)
        with upcoming_card:
            st.markdown("### Upcoming Project")
            upcoming_meetings = get_upcoming_meetings(meetings, limit=20, sort_order="Earliest deadline")
            if not upcoming_meetings:
                st.info("No upcoming projects found.")
            else:
                for meeting in upcoming_meetings[:8]:
                    with st.expander(normalize_value(meeting.get("title"), "Untitled")):
                        st.markdown(f"**Summary:** {normalize_value(meeting.get('summary'), 'No summary available.')}")
                        if meeting.get("actions"):
                            st.markdown("**Action Items**")
                            for action in meeting.get("actions", []):
                                render_action_card(action)

    with dashboard_right:
        calendar_card = st.container(border=True)
        with calendar_card:
            st.markdown("### Calendar")
            calendar_year_options = sorted(set(dashboard_years + [date.today().year]), reverse=True)
            month_names = list(calendar.month_name)[1:]
            current_month_name = calendar.month_name[date.today().month]
            selected_calendar_month = st.selectbox("Calendar Month", month_names, index=month_names.index(st.session_state.get("calendar_month", current_month_name)), key="calendar_month", label_visibility="collapsed")
            default_calendar_year = st.session_state.get("calendar_year", date.today().year)
            year_index = calendar_year_options.index(default_calendar_year) if default_calendar_year in calendar_year_options else 0
            selected_calendar_year = st.selectbox("Calendar Year", calendar_year_options, index=year_index, key="calendar_year", label_visibility="collapsed")
            selected_calendar_month_num = month_names.index(selected_calendar_month) + 1
            st.markdown(build_calendar_html(meetings, selected_calendar_year, selected_calendar_month_num), unsafe_allow_html=True)
            st.caption("Yellow dates show pending action deadlines.")

        assistant_card = st.container(border=True)
        with assistant_card:
            st.markdown("### Chatbot")
            if not meetings:
                st.info("Save a meeting first to use the assistant.")
            else:
                render_dashboard_chat(meetings)
                if st.button("Clear Chat", key="dashboard_clear_chat", use_container_width=True):
                    st.session_state.chat_history = []
                    st.rerun()
