import sys
from pathlib import Path

import streamlit as st

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.database import build_chat_thread_label, get_chat_thread_key, get_current_chat_user_id, normalize_chat_user_id
from utils.formatters import normalize_value


def render_history() -> None:
    st.subheader("Chat History")
    history_records = list(st.session_state.get("chat_history_records", []))
    current_user_id = get_current_chat_user_id()

    st.text_input("User ID", key="chat_user_id", placeholder="Enter your ID to view your history", help="Only chats saved with the same ID will appear here.")
    if not current_user_id:
        st.info("Enter your User ID to unlock private chat history.")
        st.stop()

    history_search = st.text_input("Search history", placeholder="Search by question, answer, meeting title, or context", key="history_search")
    history_context_filter = st.selectbox("Context", ["All", "Meeting", "General"], index=0, key="history_context_filter")

    user_history = [entry for entry in history_records if normalize_chat_user_id(entry.get("user_id", "")).lower() == current_user_id.lower()]
    filtered_history = user_history
    if history_context_filter != "All":
        filtered_history = [entry for entry in filtered_history if normalize_value(entry.get("context"), "General") == history_context_filter]
    search_needle = history_search.strip().lower()
    if search_needle:
        filtered_history = [
            entry for entry in filtered_history if search_needle in " ".join([
                normalize_value(entry.get("timestamp"), ""),
                normalize_value(entry.get("question"), ""),
                normalize_value(entry.get("answer"), ""),
                normalize_value(entry.get("meeting_title"), ""),
                normalize_value(entry.get("meeting_id"), ""),
                normalize_value(entry.get("context"), ""),
            ]).lower()
        ]

    st.caption(f"Showing {len(filtered_history)} of {len(user_history)} saved chat entries for ID {current_user_id}.")
    if not filtered_history:
        st.info("No chat history found for the selected search.")
        return

    grouped_threads = {}
    for entry in filtered_history:
        thread_key = normalize_value(entry.get("thread_key"), "")
        if not thread_key:
            thread_key = get_chat_thread_key(current_user_id, normalize_value(entry.get("thread_date"), normalize_value(entry.get("timestamp"), "")[:10]), normalize_value(entry.get("thread_title") or entry.get("meeting_title"), "General"), normalize_value(entry.get("meeting_id"), ""))
        grouped_threads.setdefault(thread_key, []).append(entry)

    sorted_threads = sorted(grouped_threads.items(), key=lambda item: max(normalize_value(entry.get("timestamp"), "") for entry in item[1]), reverse=True)
    for _, entries in sorted_threads:
        entries = sorted(entries, key=lambda item: normalize_value(item.get("timestamp"), ""))
        thread_head = entries[-1]
        with st.expander(build_chat_thread_label(thread_head)):
            for entry in entries:
                st.markdown(
                    f"""
                    <div class="section-card">
                        <div class="mini-title">Question</div>
                        <div class="mini-copy">{normalize_value(entry.get('question'), 'Not stated')}</div>
                        <div class="mini-title" style="margin-top:0.55rem;">Answer</div>
                        <div class="mini-copy">{normalize_value(entry.get('answer'), 'No answer saved.')}</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
