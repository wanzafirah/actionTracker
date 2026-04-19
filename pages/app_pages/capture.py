import sys
from pathlib import Path

import streamlit as st

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.database import build_email_preview_meeting, build_meeting_email_body, build_meeting_email_subject, build_meeting_record, clear_capture_inputs, clear_generated_activity_id, parse_lines, persist_app_data, set_generated_activity_id
from core.pipeline import run_pipeline
from core.services import extract_text_from_document, transcribe_audio_file
from ui.components import render_action_card, render_summary_panel
from utils.helpers import append_document_to_transcript, extract_entity_names, generate_activity_id, today_str
from utils.formatters import normalize_value


def _render_email_copy_block(meeting: dict, key_prefix: str) -> None:
    st.markdown(
        """
        <div class="section-card">
            <div class="mini-title">Email Copy Template</div>
            <div class="mini-copy">Fill in the names below, then copy the prepared subject and email body.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    recipient_name = st.text_input("Recipient Name", key=f"{key_prefix}_recipient_name", placeholder="Recipient name")
    sender_name = st.text_input("Sender Name", key=f"{key_prefix}_sender_name", placeholder="Your name")
    st.text_input("Email Subject", value=build_meeting_email_subject(meeting), key=f"{key_prefix}_subject")
    st.text_area("Email Body", value=build_meeting_email_body(meeting, recipient_name=recipient_name, sender_name=sender_name), height=260, key=f"{key_prefix}_body")


def render_capture(meetings: list, departments: list) -> None:
    st.subheader("Capture & Analyze")
    st.caption("Paste notes, upload audio, or record a meeting to generate a structured executive brief.")

    activity_box = st.container(border=True)
    with activity_box:
        act_left, act_right = st.columns(2)
        dept_names = [department["name"] for department in departments]
        with act_left:
            activity_category = st.selectbox("Category", ["Internal Meeting", "External Meeting", "Workshop"], key="capture_activity_category")
            st.text_input("Activity ID", key="capture_activity_id", placeholder="Generate or enter activity ID")
            st.button("Generate Activity ID", key="generate_activity_id_btn", on_click=set_generated_activity_id)
            st.button("Clear Activity ID", key="clear_activity_id_btn", on_click=clear_generated_activity_id)
            activity_title = st.text_input("Title", key="capture_activity_title", placeholder="Enter meeting or activity title")
            activity_type = st.selectbox("Activity Type", ["None", "Virtual", "Physical", "Both"], key="capture_activity_type")
        with act_right:
            meeting_date = st.date_input("Meeting Date", value=st.session_state.get("capture_meeting_date", today_str()), key="capture_meeting_date")
            dept_choice = st.multiselect("Department", dept_names, key="capture_department_choices")
            organization_type = st.selectbox("Organization", ["Institution", "Company"], key="capture_organization_type")
            report_by = st.text_input("Report By", key="capture_updated_by", placeholder="Enter reporter name")
            stakeholder_text = st.text_area("Stakeholders", key="capture_stakeholders", height=110, placeholder="List one stakeholder per line")

    capture_stakeholders = parse_lines(stakeholder_text)
    transcript_box = st.container(border=True)
    with transcript_box:
        audio_mode = st.radio("Audio source", ["Manual transcript", "Upload audio file", "Record meeting audio"], horizontal=True)
        transcript_mode = st.selectbox("Transcript output", ["Translate to English", "Keep spoken language"])
        document_files = st.file_uploader("Add supporting files", type=["pdf", "docx", "xlsx", "xls", "csv"], accept_multiple_files=True)

        uploaded_audio = None
        recorded_audio = None
        if audio_mode == "Upload audio file":
            uploaded_audio = st.file_uploader("Upload audio", type=["mp3", "m4a", "wav", "mp4", "mpeg", "mpga", "webm"])
        elif audio_mode == "Record meeting audio" and hasattr(st, "audio_input"):
            recorded_audio = st.audio_input("Record meeting")

        if st.button("Transcribe Audio", disabled=audio_mode == "Manual transcript" or (uploaded_audio is None and recorded_audio is None)):
            audio_source = uploaded_audio if uploaded_audio is not None else recorded_audio
            with st.spinner("Transcribing audio with local Whisper..."):
                st.session_state.capture_transcript = transcribe_audio_file(audio_source, translate_to_english=transcript_mode == "Translate to English")
                st.success("Transcript ready. Review it below before generating the summary.")

        if document_files and st.button("Add File Content"):
            for document_file in document_files:
                extracted_text = extract_text_from_document(document_file)
                labeled_text = f"File: {getattr(document_file, 'name', 'document')}\n{extracted_text}"
                st.session_state.capture_transcript = append_document_to_transcript(st.session_state.capture_transcript, labeled_text)
            st.success("Supporting document content added to the transcript area.")

        transcript = st.text_area("Transcript / Meeting Notes", height=260, key="capture_transcript")
        left, right = st.columns(2)
        with left:
            run_clicked = st.button("Generate Summary", type="primary", use_container_width=True, disabled=not transcript.strip())
        with right:
            st.button("Clear Input", on_click=clear_capture_inputs, use_container_width=True)

    resolved_activity_id = st.session_state.capture_activity_id.strip() or generate_activity_id(activity_category, meeting_date, meetings)

    if run_clicked:
        progress = st.progress(0, text="Starting pipeline...")
        progress.progress(0.25, text="Reading transcript...")
        progress.progress(0.5, text="Calling Ollama 3.2...")
        progress.progress(0.8, text="Preparing meeting brief...")
        attach_file_value = " | ".join(getattr(document_file, "name", "").strip() for document_file in (document_files or []) if getattr(document_file, "name", "").strip())
        pipeline_metadata = {
            "Category": activity_category,
            "Activity ID": resolved_activity_id,
            "Activity Title": activity_title,
            "Activity Type": activity_type,
            "Organization Type": organization_type,
            "Department": ", ".join(dept_choice),
            "Stakeholders": ", ".join(capture_stakeholders),
            "Updated By": report_by,
            "Attach File": attach_file_value,
            "Meeting Date": meeting_date.isoformat(),
        }
        result = run_pipeline(transcript, pipeline_metadata)
        progress.progress(1.0, text="Done")
        st.session_state.pending_result = {
            "result": result,
            "transcript": transcript,
            "mtype": activity_type,
            "category": activity_category,
            "dept": dept_choice,
            "meeting_date": meeting_date.isoformat(),
            "activity_category": activity_category,
            "activity_id": resolved_activity_id,
            "activity_title": activity_title,
            "organization_type": organization_type,
            "capture_stakeholders": capture_stakeholders,
            "updated_by": report_by,
            "attach_file": attach_file_value,
            "activity_type": activity_type,
        }

    if st.session_state.pending_result:
        pending = st.session_state.pending_result
        result = dict(pending["result"])
        result["title"] = result.get("title") or pending.get("activity_title", "Untitled")
        if "preview_actions" not in pending:
            pending["preview_actions"] = [
                {
                    "id": f"preview_{index}",
                    "text": normalize_value(action.get("text"), "Untitled action"),
                    "owner": normalize_value(action.get("owner"), "Not stated"),
                    "department": normalize_value(action.get("department") or action.get("company"), "Not stated"),
                    "deadline": normalize_value(action.get("deadline"), "None"),
                    "priority": action.get("priority", "Medium"),
                    "status": "Pending",
                    "suggestion": normalize_value(action.get("suggestion"), "No next-step suggestion generated."),
                }
                for index, action in enumerate(result.get("action_items", []))
            ]
        render_summary_panel(result)

        entities = result.get("nlp_pipeline", {}).get("named_entities", {})
        st.markdown("### Meeting Metadata")
        st.text_area("People", value="\n".join(extract_entity_names(entities.get("persons", []))), key="preview_people", height=100)
        st.text_area("Organizations", value="\n".join(extract_entity_names(entities.get("organizations", []))), key="preview_organizations", height=100)
        st.text_area("Dates Mentioned", value="\n".join(extract_entity_names(entities.get("dates", []))), key="preview_dates", height=100)

        st.markdown("### Action Plan")
        preview_actions = pending.get("preview_actions", [])
        if preview_actions:
            for action in preview_actions:
                render_action_card(action, editable=True)
        else:
            st.info("No action items detected.")

        email_preview_meeting = build_email_preview_meeting(result, pending)
        with st.expander("Prepare Email Copy"):
            _render_email_copy_block(email_preview_meeting, "preview_email_copy")

        if st.button("Save Record", type="primary"):
            new_meeting = build_meeting_record(result, pending)
            st.session_state.meetings.insert(0, new_meeting)
            persist_app_data()
            st.session_state.pending_result = None
            st.success("Meeting saved to library.")
            st.rerun()
