import sys
from pathlib import Path

import streamlit as st

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.database import parse_lines, persist_app_data, update_meeting_record
from ui.components import render_action_card, render_completion_ring, render_kpi_card
from utils.formatters import normalize_status, normalize_value


def render_tracker(meetings: list) -> None:
    st.subheader("Action Tracker")

    def meeting_tracker_status(meeting: dict) -> str:
        actions = meeting.get("actions", [])
        if not actions:
            return "Completed"
        statuses = [normalize_status(action) for action in actions]
        if "Overdue" in statuses:
            return "Overdue"
        if any(status in {"Pending", "In Progress"} for status in statuses):
            return "Pending"
        return "Completed"

    tracker_focus = st.session_state.get("tracker_focus", "all")
    status_default = "All"
    if tracker_focus == "open":
        status_default = "Pending"
    elif tracker_focus == "done":
        status_default = "Completed"

    meeting_records = []
    for meeting in meetings:
        meeting_records.append(
            {
                "meeting": meeting,
                "status": meeting_tracker_status(meeting),
                "meeting_id": normalize_value(meeting.get("meetingID") or meeting.get("activityId") or meeting.get("id"), ""),
                "title": normalize_value(meeting.get("title"), "Untitled meeting"),
                "group": normalize_value(meeting.get("deptName") or meeting.get("department") or meeting.get("sltdepartment"), ""),
                "summary": normalize_value(meeting.get("summary") or meeting.get("recaps"), ""),
                "keywords": " ".join([
                    normalize_value(meeting.get("title"), ""),
                    normalize_value(meeting.get("summary"), ""),
                    normalize_value(meeting.get("meetingID"), ""),
                    normalize_value(meeting.get("department"), ""),
                ]).lower(),
            }
        )

    tracker_actions = [action for meeting in meetings for action in meeting.get("actions", [])]
    tracker_statuses = [normalize_status(action) for action in tracker_actions]
    total_tracker_actions = len(tracker_actions)
    pending_tracker_actions = sum(1 for status in tracker_statuses if status in {"Pending", "In Progress", "Overdue"})
    completed_tracker_actions = sum(1 for status in tracker_statuses if status == "Done")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        render_kpi_card("Total Meeting", str(len(meeting_records)), "Stored records", "#0f766e")
    with c2:
        render_kpi_card("Total Action Item", str(total_tracker_actions), "Tracked tasks", "#1e3a5f")
    with c3:
        render_kpi_card("Follow up needed", str(pending_tracker_actions), "Open tasks", "#d97706")
    with c4:
        completion_pct = round((completed_tracker_actions / total_tracker_actions) * 100) if total_tracker_actions else 0
        render_completion_ring(completion_pct)

    meeting_search = st.text_input("Search meeting", placeholder="Search by title, ID, or department", key="tracker_meeting_search")
    status_options = ["All", "Pending", "Overdue", "Completed"]
    status_index = status_options.index(status_default) if status_default in status_options else 0
    meeting_status_filter = st.selectbox("Meeting Status", status_options, index=status_index, key="tracker_meeting_status")

    filtered_meetings = meeting_records
    if meeting_status_filter != "All":
        filtered_meetings = [record for record in filtered_meetings if record["status"] == meeting_status_filter]
    search_needle = meeting_search.strip().lower()
    if search_needle:
        filtered_meetings = [record for record in filtered_meetings if search_needle in record["keywords"]]

    st.markdown("### Saved Meetings")
    if not filtered_meetings:
        st.info("No saved meetings match the selected search or status.")
        return

    filtered_meetings = sorted(filtered_meetings, key=lambda record: normalize_value(record["meeting"].get("date"), "0000-00-00"), reverse=True)
    for record in filtered_meetings:
        meeting = record["meeting"]
        meeting_id_value = normalize_value(meeting.get("meetingID") or meeting.get("activityId") or meeting.get("id"), "")
        header = f"{record['title']} | {record['meeting_id'] or 'No ID'} | {normalize_value(meeting.get('date'), 'No date')}"
        with st.expander(header):
            st.markdown(f"**Summary:** {normalize_value(meeting.get('summary'), 'No summary available.')}")
            if st.button("Edit Meeting", key=f"edit_meeting_btn_{meeting_id_value}"):
                st.session_state.tracker_editing_meeting_id = meeting_id_value
            if st.session_state.get("tracker_editing_meeting_id") == meeting_id_value:
                with st.form(key=f"meeting_edit_form_{meeting_id_value}", clear_on_submit=False):
                    edit_title = st.text_input("Title", value=normalize_value(meeting.get("title"), ""), key=f"edit_title_{meeting_id_value}")
                    edit_date = st.text_input("Date", value=normalize_value(meeting.get("date"), ""), key=f"edit_date_{meeting_id_value}")
                    edit_department = st.text_input("Group / Department", value=normalize_value(meeting.get("deptName") or meeting.get("department") or meeting.get("sltdepartment"), ""), key=f"edit_department_{meeting_id_value}")
                    edit_summary = st.text_area("Summary", value=normalize_value(meeting.get("summary") or meeting.get("recaps"), ""), key=f"edit_summary_{meeting_id_value}")
                    edit_objective = st.text_area("Objective", value=normalize_value(meeting.get("objective"), ""), key=f"edit_objective_{meeting_id_value}")
                    edit_outcome = st.text_area("Outcome", value=normalize_value(meeting.get("outcome"), ""), key=f"edit_outcome_{meeting_id_value}")
                    edit_stakeholders = st.text_area("Stakeholders", value="\n".join(meeting.get("stakeholders", []) or []), key=f"edit_stakeholders_{meeting_id_value}")
                    edit_companies = st.text_area("Companies", value="\n".join(meeting.get("companies", []) or []), key=f"edit_companies_{meeting_id_value}")
                    save_edit = st.form_submit_button("Save Changes")
                if save_edit:
                    meeting_updates = {
                        "title": edit_title.strip() or normalize_value(meeting.get("title"), "Untitled meeting"),
                        "date": edit_date.strip() or normalize_value(meeting.get("date"), ""),
                        "meeting date": edit_date.strip() or normalize_value(meeting.get("date"), ""),
                        "deptName": edit_department.strip(),
                        "department": edit_department.strip(),
                        "sltdepartment": edit_department.strip(),
                        "summary": edit_summary.strip(),
                        "recaps": edit_summary.strip(),
                        "objective": edit_objective.strip(),
                        "outcome": edit_outcome.strip(),
                        "stakeholders": parse_lines(edit_stakeholders),
                        "companies": parse_lines(edit_companies),
                        "updatedBy": "Manual edit",
                    }
                    if update_meeting_record(st.session_state.meetings, meeting_id_value, meeting_updates):
                        persist_app_data()
                        st.session_state.tracker_editing_meeting_id = ""
                        st.success("Meeting updated successfully.")
                        st.rerun()
            if meeting.get("actions"):
                st.markdown("**Action Items**")
                for action in meeting["actions"]:
                    render_action_card(action, editable=True, persist_callback=persist_app_data)
            else:
                st.success("No action item. This meeting is considered completed.")
