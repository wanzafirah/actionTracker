import json
import os
import re
import calendar
import ast
from datetime import date, datetime

import pandas as pd
import plotly.express as px
import requests
import streamlit as st

try:
    import gspread
except ImportError:
    gspread = None

from meetiq_constants import (
    ACTIVITY_CATEGORY_OPTIONS,
    ACTIVITY_TYPE_OPTIONS,
    CATEGORIES,
    DATA_FILE,
    DEFAULT_DEPARTMENTS,
    DEPARTMENTS_SHEET_NAME,
    DEPARTMENT_SHEET_COLUMNS,
    LINK_PHOTO_OPTIONS,
    MAIN_ACTIVITY_OPTIONS,
    MEETINGS_SHEET_NAME,
    MEETING_SHEET_COLUMNS,
    MTG_TYPES,
    ORGANIZATION_TYPE_OPTIONS,
    REPRESENTATIVE_POSITION_OPTIONS,
    ROLE_OPTIONS,
    STATUSES,
)
from meetiq_services import call_ollama, extract_text_from_document, transcribe_audio_file
from meetiq_ui import render_action_card, render_chat_bubble, render_completion_ring, render_kpi_card, render_summary_panel
from meetiq_utils import (
    action_belongs_to_talentcorp,
    add_month_columns,
    append_document_to_transcript,
    build_action_dataframe,
    build_calendar_html,
    build_meeting_dataframe,
    compact_transcript_for_prompt,
    days_left,
    entity_text,
    extract_entity_names,
    fallback_discussion_points,
    fallback_key_decisions,
    fallback_action_items,
    filter_talentcorp_actions,
    first_nonempty,
    generate_activity_id,
    get_pending_deadline_days,
    get_upcoming_meetings,
    html_lines,
    is_objective_only_transcript,
    join_list,
    json_dumps_safe,
    json_loads_safe,
    load_text_list,
    normalize_status,
    normalize_value,
    parse_yes_no,
    pretty_deadline,
    render_entity_list,
    render_plotly_chart,
    rm,
    style_plotly,
    today_str,
    uid,
    yes_no_text,
)


def get_secret_value(name: str, default=""):
    try:
        value = st.secrets.get(name, default)
        return value if value not in (None, "") else default
    except Exception:
        return default


def get_google_service_account_info():
    try:
        secret_value = st.secrets.get("gcp_service_account")
        if secret_value:
            return dict(secret_value)
    except Exception:
        pass

    raw_json = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "")
    if raw_json:
        try:
            return json.loads(raw_json)
        except json.JSONDecodeError:
            return None
    return None


def get_google_sheet_target() -> dict:
    return {
        "id": os.getenv("GOOGLE_SHEETS_SPREADSHEET_ID", get_secret_value("GOOGLE_SHEETS_SPREADSHEET_ID", "")),
        "url": os.getenv("GOOGLE_SHEETS_SPREADSHEET_URL", get_secret_value("GOOGLE_SHEETS_SPREADSHEET_URL", "")),
        "name": os.getenv("GOOGLE_SHEETS_SPREADSHEET_NAME", get_secret_value("GOOGLE_SHEETS_SPREADSHEET_NAME", "")),
    }


def get_supabase_config() -> dict:
    url = os.getenv("SUPABASE_URL", get_secret_value("SUPABASE_URL", ""))
    key = (
        os.getenv("SUPABASE_SERVICE_ROLE_KEY", get_secret_value("SUPABASE_SERVICE_ROLE_KEY", ""))
        or os.getenv("SUPABASE_KEY", get_secret_value("SUPABASE_KEY", ""))
        or os.getenv("SUPABASE_ANON_KEY", get_secret_value("SUPABASE_ANON_KEY", ""))
    )
    return {"url": url.rstrip("/"), "key": key}


def set_generated_activity_id():
    category = st.session_state.get("capture_activity_category", "")
    meeting_date_value = st.session_state.get("capture_meeting_date", date.today())
    st.session_state.capture_activity_id = generate_activity_id(category, meeting_date_value, st.session_state.get("meetings", []))


def clear_generated_activity_id():
    st.session_state.capture_activity_id = ""


def set_current_page(page_name: str):
    st.session_state.current_page = page_name


def set_tracker_shortcut(mode: str):
    st.session_state.current_page = "Tracker"
    st.session_state.tracker_focus = mode


def sync_page_from_query():
    try:
        page_value = st.query_params.get("page")
        focus_value = st.query_params.get("focus")
    except Exception:
        return

    if page_value in {"Dashboard", "Productivity", "Capture", "Tracker", "Finance"}:
        st.session_state.current_page = page_value
    if focus_value in {"all", "open", "done"}:
        st.session_state.tracker_focus = focus_value


def clear_capture_inputs():
    st.session_state.pending_result = None
    st.session_state.capture_transcript = ""
    st.session_state.capture_activity_id = ""
    st.session_state.capture_link_photo_url = ""
    st.session_state.capture_district = ""
    st.session_state.capture_invitation_from = ""
    st.session_state.capture_location_meeting = ""
    st.session_state.capture_other_reps = ""
    st.session_state.capture_stfemail = ""
    st.session_state.capture_supemail = ""
    st.session_state.capture_updated_by = ""


def parse_loaded_dataframes(meetings_df: pd.DataFrame, departments_df: pd.DataFrame):
    meetings = []
    if not meetings_df.empty:
        for row in meetings_df.fillna("").to_dict("records"):
            meetings.append(
                {
                    "id": first_nonempty(row, "id", fallback=""),
                    "title": first_nonempty(row, "title", "activityTitle", fallback=""),
                    "date": first_nonempty(row, "meeting date", "date", "meetingDate", fallback=""),
                    "type": first_nonempty(row, "meeting type", "type", "activityType", fallback=""),
                    "category": first_nonempty(row, "category", "activityCategory", fallback=""),
                    "summary": first_nonempty(row, "recaps", "summary", fallback=""),
                    "objective": first_nonempty(row, "objective", "activityObjective", fallback=""),
                    "outcome": first_nonempty(row, "outcome", fallback=""),
                    "followUp": parse_yes_no(first_nonempty(row, "followup", "followUp", fallback=False)),
                    "followUpReason": first_nonempty(row, "followUpReason", fallback=""),
                    "stakeholders": load_text_list(first_nonempty(row, "stakeholders", fallback=[])),
                    "companies": load_text_list(first_nonempty(row, "companies", fallback=[])),
                    "keyDecisions": json_loads_safe(first_nonempty(row, "keyDecisions", fallback=[]), []),
                    "discussionPoints": json_loads_safe(first_nonempty(row, "discussionPoints", fallback=[]), []),
                    "nlpStats": json_loads_safe(first_nonempty(row, "nlpStats", fallback={}), {}),
                    "transcript": first_nonempty(row, "transcript", "recaps", fallback=""),
                    "deptId": first_nonempty(row, "deptId", fallback=""),
                    "deptName": first_nonempty(row, "deptName", "department", "sltdepartment", fallback=""),
                    "department": first_nonempty(row, "department", "deptName", "sltdepartment", fallback=""),
                    "actualCost": float(first_nonempty(row, "actualCost", "budgetUsed", fallback=0) or 0),
                    "budgetUsed": float(first_nonempty(row, "budgetUsed", "actualCost", fallback=0) or 0),
                    "estimatedCost": float(first_nonempty(row, "estimatedCost", fallback=0) or 0),
                    "budgetNotes": first_nonempty(row, "budgetNotes", fallback=""),
                    "actions": json_loads_safe(first_nonempty(row, "actions", fallback=[]), []),
                    "activityCategory": first_nonempty(row, "activityCategory", "category", fallback=""),
                    "activityId": first_nonempty(row, "meetingID", "activityId", "id", fallback=""),
                    "activityTitle": first_nonempty(row, "activityTitle", "title", fallback=""),
                    "role": first_nonempty(row, "role", fallback=""),
                    "mainActivity": first_nonempty(row, "mainActivity", fallback=""),
                    "linkPhoto": first_nonempty(row, "linkPhoto", "attach file", fallback=""),
                    "linkPhotoUrl": first_nonempty(row, "linkPhotoUrl", "attach file", fallback=""),
                    "activityType": first_nonempty(row, "activityType", "meeting type", "type", fallback=""),
                    "organizationType": first_nonempty(row, "organizationType", fallback=""),
                    "dateFrom": first_nonempty(row, "dateFrom", fallback=""),
                    "dateTo": first_nonempty(row, "dateTo", fallback=""),
                    "representativePosition": first_nonempty(row, "representativePosition", "sltposition", fallback=""),
                    "representativeName": first_nonempty(row, "representativeName", "sltreps", fallback=""),
                    "representativeDepartment": first_nonempty(row, "representativeDepartment", "sltdepartment", fallback=""),
                    "activityObjective": first_nonempty(row, "activityObjective", fallback=""),
                    "attachFile": first_nonempty(row, "attach file", "linkPhotoUrl", "linkPhoto", fallback=""),
                    "district": first_nonempty(row, "district", fallback=""),
                    "invitationFrom": first_nonempty(row, "invitationfrom", fallback=""),
                    "locationMeeting": first_nonempty(row, "location meeting", fallback=""),
                    "meetingID": first_nonempty(row, "meetingID", "activityId", "id", fallback=""),
                    "meetingType": first_nonempty(row, "meeting type", "activityType", "type", fallback=""),
                    "otherReps": first_nonempty(row, "other reps", fallback=""),
                    "recaps": first_nonempty(row, "recaps", "summary", fallback=""),
                    "sltdepartment": first_nonempty(row, "sltdepartment", "deptName", "department", fallback=""),
                    "sltposition": first_nonempty(row, "sltposition", "representativePosition", fallback=""),
                    "sltreps": first_nonempty(row, "sltreps", "representativeName", fallback=""),
                    "stfemail": first_nonempty(row, "stfemail", fallback=""),
                    "supemail": first_nonempty(row, "supemail", fallback=""),
                    "updatedBy": first_nonempty(row, "updated by", fallback=""),
                }
            )

    departments = []
    if not departments_df.empty:
        for row in departments_df.fillna("").to_dict("records"):
            departments.append(
                {
                    "id": row.get("id", ""),
                    "name": row.get("name", ""),
                    "budget": float(row.get("budget", 0) or 0),
                }
            )

    return meetings, departments


def build_meeting_rows(meetings: list) -> list:
    meetings_rows = []
    for meeting in meetings:
        attach_file_value = normalize_value(
            meeting.get("attachFile") or meeting.get("linkPhotoUrl") or meeting.get("linkPhoto"),
            "",
        )
        meetings_rows.append(
            {
                "id": meeting.get("id", ""),
                "title": meeting.get("title", ""),
                "date": meeting.get("date", ""),
                "meeting date": meeting.get("date", ""),
                "type": meeting.get("type", ""),
                "meeting type": meeting.get("type", ""),
                "category": meeting.get("category", ""),
                "district": meeting.get("district", ""),
                "summary": meeting.get("summary", ""),
                "objective": meeting.get("objective", ""),
                "outcome": meeting.get("outcome", ""),
                "followUp": meeting.get("followUp", False),
                "followup": yes_no_text(meeting.get("followUp", False)),
                "followUpReason": meeting.get("followUpReason", ""),
                "stakeholders": json_dumps_safe(meeting.get("stakeholders", [])),
                "companies": json_dumps_safe(meeting.get("companies", [])),
                "keyDecisions": json_dumps_safe(meeting.get("keyDecisions", [])),
                "discussionPoints": json_dumps_safe(meeting.get("discussionPoints", [])),
                "nlpStats": json_dumps_safe(meeting.get("nlpStats", {})),
                "transcript": meeting.get("transcript", ""),
                "deptId": meeting.get("deptId", ""),
                "deptName": meeting.get("deptName", meeting.get("department", "")),
                "department": meeting.get("department", meeting.get("deptName", "")),
                "actualCost": meeting.get("actualCost", meeting.get("budgetUsed", 0)),
                "budgetUsed": meeting.get("budgetUsed", meeting.get("actualCost", 0)),
                "estimatedCost": meeting.get("estimatedCost", 0),
                "budgetNotes": meeting.get("budgetNotes", ""),
                "actions": json_dumps_safe(meeting.get("actions", [])),
                "activityCategory": meeting.get("activityCategory", ""),
                "activityId": meeting.get("activityId", ""),
                "meetingID": meeting.get("meetingID", meeting.get("activityId", meeting.get("id", ""))),
                "activityTitle": meeting.get("activityTitle", ""),
                "role": meeting.get("role", ""),
                "mainActivity": meeting.get("mainActivity", ""),
                "linkPhoto": meeting.get("linkPhoto", ""),
                "linkPhotoUrl": meeting.get("linkPhotoUrl", ""),
                "attach file": attach_file_value,
                "activityType": meeting.get("activityType", ""),
                "organizationType": meeting.get("organizationType", ""),
                "dateFrom": meeting.get("dateFrom", ""),
                "dateTo": meeting.get("dateTo", ""),
                "representativePosition": meeting.get("representativePosition", ""),
                "representativeName": meeting.get("representativeName", ""),
                "representativeDepartment": meeting.get("representativeDepartment", ""),
                "activityObjective": meeting.get("activityObjective", ""),
                "invitationfrom": meeting.get("invitationFrom", ""),
                "location meeting": meeting.get("locationMeeting", ""),
                "other reps": meeting.get("otherReps", ""),
                "recaps": meeting.get("recaps", meeting.get("summary", "")),
                "sltdepartment": meeting.get("sltdepartment", meeting.get("deptName", meeting.get("department", ""))),
                "sltposition": meeting.get("sltposition", meeting.get("representativePosition", "")),
                "sltreps": meeting.get("sltreps", meeting.get("representativeName", "")),
                "stfemail": meeting.get("stfemail", ""),
                "supemail": meeting.get("supemail", ""),
                "updated by": meeting.get("updatedBy", ""),
            }
        )
    return meetings_rows


def build_department_rows(departments: list) -> list:
    return [
        {
            "id": department.get("id", ""),
            "name": department.get("name", ""),
            "budget": department.get("budget", 0),
        }
        for department in departments
    ]


def load_excel_data():
    if not os.path.exists(DATA_FILE):
        return [], []

    try:
        meetings_df = pd.read_excel(DATA_FILE, sheet_name=MEETINGS_SHEET_NAME)
    except Exception:
        meetings_df = pd.DataFrame()

    try:
        departments_df = pd.read_excel(DATA_FILE, sheet_name=DEPARTMENTS_SHEET_NAME)
    except Exception:
        departments_df = pd.DataFrame()

    return parse_loaded_dataframes(meetings_df, departments_df)


def save_excel_data(meetings: list, departments: list):
    meetings_rows = build_meeting_rows(meetings)
    departments_rows = build_department_rows(departments)

    with pd.ExcelWriter(DATA_FILE, engine="openpyxl") as writer:
        pd.DataFrame(meetings_rows, columns=MEETING_SHEET_COLUMNS).to_excel(writer, sheet_name=MEETINGS_SHEET_NAME, index=False)
        pd.DataFrame(departments_rows, columns=DEPARTMENT_SHEET_COLUMNS).to_excel(writer, sheet_name=DEPARTMENTS_SHEET_NAME, index=False)


@st.cache_resource
def get_google_spreadsheet():
    service_account_info = get_google_service_account_info()
    target = get_google_sheet_target()
    if gspread is None or not service_account_info:
        return None

    if not any(target.values()):
        return None

    client = gspread.service_account_from_dict(service_account_info)
    if target["id"]:
        return client.open_by_key(target["id"])
    if target["url"]:
        return client.open_by_url(target["url"])
    return client.open(target["name"])


def get_or_create_google_worksheet(spreadsheet, title: str, columns: list):
    try:
        worksheet = spreadsheet.worksheet(title)
    except Exception:
        worksheet = spreadsheet.add_worksheet(title=title, rows=max(100, len(columns) + 20), cols=len(columns))

    header = worksheet.row_values(1)
    if header != columns:
        worksheet.clear()
        worksheet.update("A1", [columns])
    return worksheet


def load_google_sheet_data():
    spreadsheet = get_google_spreadsheet()
    if spreadsheet is None:
        return None

    try:
        meetings_ws = get_or_create_google_worksheet(spreadsheet, MEETINGS_SHEET_NAME, MEETING_SHEET_COLUMNS)
        departments_ws = get_or_create_google_worksheet(spreadsheet, DEPARTMENTS_SHEET_NAME, DEPARTMENT_SHEET_COLUMNS)
        meetings_df = pd.DataFrame(meetings_ws.get_all_records())
        departments_df = pd.DataFrame(departments_ws.get_all_records())
        return parse_loaded_dataframes(meetings_df, departments_df)
    except Exception:
        return None


def save_google_sheet_data(meetings: list, departments: list):
    spreadsheet = get_google_spreadsheet()
    if spreadsheet is None:
        raise RuntimeError(
            "Google Sheets storage is not configured. Add gcp_service_account and a spreadsheet ID, URL, or name in Streamlit secrets."
        )

    meetings_rows = build_meeting_rows(meetings)
    departments_rows = build_department_rows(departments)

    meetings_ws = get_or_create_google_worksheet(spreadsheet, MEETINGS_SHEET_NAME, MEETING_SHEET_COLUMNS)
    meetings_values = [MEETING_SHEET_COLUMNS] + [
        [row.get(column, "") for column in MEETING_SHEET_COLUMNS]
        for row in meetings_rows
    ]
    meetings_ws.clear()
    meetings_ws.update("A1", meetings_values)

    departments_ws = get_or_create_google_worksheet(spreadsheet, DEPARTMENTS_SHEET_NAME, DEPARTMENT_SHEET_COLUMNS)
    departments_values = [DEPARTMENT_SHEET_COLUMNS] + [
        [row.get(column, "") for column in DEPARTMENT_SHEET_COLUMNS]
        for row in departments_rows
    ]
    departments_ws.clear()
    departments_ws.update("A1", departments_values)


def supabase_headers(prefer: str = "") -> dict | None:
    config = get_supabase_config()
    if not config["url"] or not config["key"]:
        return None

    headers = {
        "apikey": config["key"],
        "Authorization": f"Bearer {config['key']}",
        "Content-Type": "application/json",
    }
    if prefer:
        headers["Prefer"] = prefer
    return headers


def supabase_table_url(table_name: str) -> str | None:
    config = get_supabase_config()
    if not config["url"]:
        return None
    return f"{config['url']}/rest/v1/{table_name}"


def load_supabase_data():
    headers = supabase_headers()
    meetings_url = supabase_table_url(MEETINGS_SHEET_NAME)
    departments_url = supabase_table_url(DEPARTMENTS_SHEET_NAME)
    if headers is None or meetings_url is None or departments_url is None:
        return None

    try:
        meetings_response = requests.get(meetings_url, headers=headers, params={"select": "*"}, timeout=30)
        departments_response = requests.get(departments_url, headers=headers, params={"select": "*"}, timeout=30)
        meetings_response.raise_for_status()
        departments_response.raise_for_status()
        meetings_df = pd.DataFrame(meetings_response.json())
        departments_df = pd.DataFrame(departments_response.json())
        return parse_loaded_dataframes(meetings_df, departments_df)
    except Exception:
        return None


def replace_supabase_table(table_name: str, rows: list, columns: list):
    headers = supabase_headers(prefer="return=minimal")
    table_url = supabase_table_url(table_name)
    if headers is None or table_url is None:
        raise RuntimeError("Supabase storage is not configured.")

    delete_response = requests.delete(table_url, headers=headers, params={"id": "neq.__keep__"}, timeout=30)
    delete_response.raise_for_status()

    if not rows:
        return

    payload = [{column: row.get(column, "") for column in columns} for row in rows]
    insert_response = requests.post(table_url, headers=headers, json=payload, timeout=30)
    insert_response.raise_for_status()


def save_supabase_data(meetings: list, departments: list):
    meeting_rows = build_meeting_rows(meetings)
    department_rows = build_department_rows(departments)
    replace_supabase_table(MEETINGS_SHEET_NAME, meeting_rows, MEETING_SHEET_COLUMNS)
    replace_supabase_table(DEPARTMENTS_SHEET_NAME, department_rows, DEPARTMENT_SHEET_COLUMNS)


def load_app_data():
    supabase_data = load_supabase_data()
    if supabase_data is not None:
        return supabase_data

    google_data = load_google_sheet_data()
    if google_data is not None:
        return google_data
    return load_excel_data()


def save_app_data(meetings: list, departments: list):
    if supabase_headers() is not None:
        save_supabase_data(meetings, departments)
    elif get_google_spreadsheet() is not None:
        save_google_sheet_data(meetings, departments)
    else:
        save_excel_data(meetings, departments)


def persist_app_data():
    save_app_data(st.session_state.meetings, st.session_state.departments)


def seed_default_departments():
    existing_names = {department.get("name", "").strip().lower() for department in st.session_state.departments}
    added = False
    for department_name in DEFAULT_DEPARTMENTS:
        if department_name.strip().lower() not in existing_names:
            st.session_state.departments.append({"id": uid(), "name": department_name, "budget": 0.0})
            added = True
    if added:
        persist_app_data()


def get_department_options() -> list:
    return ["— None —"] + [department["name"] for department in st.session_state.departments]


def find_department_by_name(name: str):
    return next((department for department in st.session_state.departments if department["name"] == name), None)


def append_document_to_transcript(current_text: str, extracted_text: str) -> str:
    current_text = current_text.strip()
    if current_text:
        return f"{current_text}\n\nSupporting document:\n{extracted_text}"
    return extracted_text


def compact_transcript_for_prompt(text: str, max_chars: int = 1200) -> str:
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if not lines:
        return text[:max_chars]

    keywords = (
        "date",
        "deadline",
        "follow",
        "action",
        "decision",
        "objective",
        "month",
        "launch",
        "program",
        "task",
        "assign",
        "confirm",
        "send",
        "review",
        "recap",
        "summary",
        "deadline",
        "owner",
        "assignee",
    )

    kept = []
    for line in lines:
        lower = line.lower()
        if any(keyword in lower for keyword in keywords):
            kept.append(line)
        elif len(kept) < 30:
            kept.append(line)
        if len("\n".join(kept)) >= max_chars:
            break

    compacted = "\n".join(kept)
    return compacted[:max_chars] if compacted else text[:max_chars]


def transcript_sentences(text: str) -> list:
    raw_sentences = re.split(r"(?<=[.!?])\s+|\n+", text)
    return [sentence.strip(" -•\t") for sentence in raw_sentences if sentence and sentence.strip()]


def fallback_discussion_points(text: str, limit: int = 4) -> list:
    points = []
    for sentence in transcript_sentences(text):
        lowered = sentence.lower()
        if any(token in lowered for token in ("discuss", "review", "share", "explore", "align", "overview", "about")):
            points.append(sentence)
        elif len(points) < limit and len(sentence.split()) >= 6:
            points.append(sentence)
        if len(points) >= limit:
            break
    return points[:limit]


def fallback_key_decisions(text: str, limit: int = 3) -> list:
    decisions = []
    markers = ("decided", "agreed", "approved", "confirmed", "will proceed", "it was decided", "the team agreed")
    for sentence in transcript_sentences(text):
        lowered = sentence.lower()
        if any(marker in lowered for marker in markers):
            decisions.append(sentence)
        if len(decisions) >= limit:
            break
    return decisions[:limit]


def is_objective_only_transcript(text: str) -> bool:
    lower = " ".join(text.lower().split())
    objective_markers = [
        "the purpose of this meeting is",
        "meeting to discuss",
        "the meeting was to",
        "the meeting is to",
        "this meeting is to",
        "to coordinate the preparations",
        "to align the program outline",
        "to explore potential collaboration",
        "aiming to",
        "the discussion will explore",
    ]
    explicit_action_markers = [
        "action item",
        "assigned",
        "responsible",
        "owner",
        "deadline",
        "needs to",
        "need to",
        "must",
        "shall",
        "will prepare",
        "will send",
        "will confirm",
        "please prepare",
        "please send",
    ]
    return any(marker in lower for marker in objective_markers) and not any(
        marker in lower for marker in explicit_action_markers
    )


def make_action_preview_item(action: dict, action_id: str, status: str = "Pending") -> dict:
    return {
        "id": action_id,
        "text": normalize_value(action.get("text"), "Untitled action"),
        "owner": normalize_value(action.get("owner"), "Not stated"),
        "company": normalize_value(action.get("company"), "Internal"),
        "deadline": normalize_value(action.get("deadline"), "None"),
        "priority": action.get("priority", "Medium"),
        "status": status,
        "suggestion": normalize_value(action.get("suggestion"), "No next-step suggestion generated."),
        "followUpRequired": action.get("follow_up_required", False),
        "followUpReason": normalize_value(action.get("follow_up_reason"), "None"),
        "nerEntities": action.get("ner_entities", []),
    }


def build_email_preview_meeting(result: dict, pending: dict) -> dict:
    preview_actions = pending.get("preview_actions")
    if preview_actions:
        actions = [
            {
                "text": normalize_value(action.get("text"), "Untitled action"),
                "owner": normalize_value(action.get("owner"), "Not stated"),
                "deadline": normalize_value(action.get("deadline"), "None"),
                "status": normalize_value(action.get("status"), "Pending"),
            }
            for action in preview_actions
        ]
    else:
        actions = [
            {
                "text": normalize_value(action.get("text"), "Untitled action"),
                "owner": normalize_value(action.get("owner"), "Not stated"),
                "deadline": normalize_value(action.get("deadline"), "None"),
                "status": "Pending",
            }
            for action in result.get("action_items", [])
        ]
    return {
        "title": result.get("title") or pending.get("activity_title", "Untitled"),
        "date": pending.get("meeting_date", today_str()),
        "type": pending["mtype"],
        "category": pending["category"],
        "summary": result.get("summary", ""),
        "objective": result.get("objective", ""),
        "outcome": result.get("outcome", ""),
        "followUp": result.get("follow_up", False),
        "followUpReason": result.get("follow_up_reason", ""),
        "keyDecisions": result.get("key_decisions", []),
        "actions": actions,
    }


def build_meeting_record(result: dict, pending: dict) -> dict:
    meeting_id = uid()
    department = find_department_by_name(pending["dept"])
    preview_actions = pending.get("preview_actions", [])
    return {
        "id": meeting_id,
        "title": result.get("title") or pending.get("activity_title", "Untitled"),
        "date": pending.get("meeting_date", today_str()),
        "meeting date": pending.get("meeting_date", today_str()),
        "type": pending["mtype"],
        "meeting type": pending["mtype"],
        "category": pending["category"],
        "district": pending.get("district", ""),
        "summary": result.get("summary", ""),
        "recaps": result.get("summary", ""),
        "objective": result.get("objective", ""),
        "outcome": result.get("outcome", ""),
        "followUp": result.get("follow_up", False),
        "followUpReason": result.get("follow_up_reason", "") or "",
        "stakeholders": extract_entity_names(result.get("nlp_pipeline", {}).get("named_entities", {}).get("persons", [])),
        "companies": extract_entity_names(result.get("nlp_pipeline", {}).get("named_entities", {}).get("organizations", [])),
        "keyDecisions": result.get("key_decisions", []),
        "discussionPoints": result.get("discussion_points", []),
        "nlpStats": result.get("nlp_pipeline", {}),
        "transcript": pending["transcript"],
        "deptId": department["id"] if department else "",
        "deptName": department["name"] if department else "",
        "department": department["name"] if department else "",
        "actualCost": pending["actual_cost"],
        "budgetUsed": pending["actual_cost"],
        "estimatedCost": result.get("estimated_budget", 0),
        "budgetNotes": result.get("budget_notes", ""),
        "activityCategory": pending.get("activity_category", ""),
        "activityId": pending.get("activity_id", ""),
        "meetingID": pending.get("activity_id", ""),
        "activityTitle": pending.get("activity_title", ""),
        "role": pending.get("role", ""),
        "mainActivity": pending.get("main_activity", ""),
        "linkPhoto": pending.get("link_photo", ""),
        "linkPhotoUrl": pending.get("link_photo_url", ""),
        "attachFile": pending.get("attach_file", ""),
        "activityType": pending.get("activity_type", ""),
        "organizationType": pending.get("organization_type", ""),
        "dateFrom": pending.get("date_from", ""),
        "dateTo": pending.get("date_to", ""),
        "representativePosition": pending.get("representative_position", ""),
        "representativeName": pending.get("representative_name", ""),
        "representativeDepartment": pending.get("representative_department", ""),
        "sltdepartment": pending.get("slt_department", ""),
        "sltposition": pending.get("slt_position", ""),
        "sltreps": pending.get("slt_reps", ""),
        "stfemail": pending.get("stf_email", ""),
        "supemail": pending.get("sup_email", ""),
        "invitationFrom": pending.get("invitation_from", ""),
        "locationMeeting": pending.get("location_meeting", ""),
        "otherReps": pending.get("other_reps", ""),
        "updatedBy": pending.get("updated_by", ""),
        "activityObjective": pending.get("activity_objective", ""),
        "actions": (
            [
                {
                    **action,
                    "id": f"{meeting_id}_a{index}",
                }
                for index, action in enumerate(preview_actions)
            ]
            if preview_actions
            else [
                make_action_preview_item(action, f"{meeting_id}_a{index}")
                for index, action in enumerate(result.get("action_items", []))
            ]
        ),
    }


def filter_meeting_library(meetings: list, search_text: str) -> list:
    filtered = []
    for meeting in meetings:
        haystack = (
            meeting["title"] + meeting["summary"] + meeting.get("outcome", "") + " ".join(meeting.get("companies", []))
        ).lower()
        if not search_text or search_text.lower() in haystack:
            filtered.append(meeting)
    return filtered


def filter_action_records(action_df: pd.DataFrame, company_filter: str, status_filter: str, action_search: str) -> pd.DataFrame:
    filtered_actions = action_df.copy()
    if company_filter != "All":
        company_needle = company_filter.lower()
        filtered_actions = filtered_actions[
            (filtered_actions["company"].str.lower() == company_needle)
            | (filtered_actions["meeting_title"].str.lower().str.contains(company_needle))
        ]
    if status_filter != "All":
        filtered_actions = filtered_actions[filtered_actions["status"] == status_filter]
    if action_search.strip():
        needle = action_search.strip().lower()
        filtered_actions = filtered_actions[
            filtered_actions["text"].str.lower().str.contains(needle)
            | filtered_actions["owner"].str.lower().str.contains(needle)
            | filtered_actions["meeting_title"].str.lower().str.contains(needle)
        ]
    return filtered_actions


def build_meeting_email_subject(meeting: dict) -> str:
    return f"Meeting Summary | {normalize_value(meeting.get('title'), 'Untitled Meeting')} | {normalize_value(meeting.get('date'), today_str())}"


def build_meeting_email_body(meeting: dict, recipient_name: str = "", sender_name: str = "") -> str:
    greeting_name = recipient_name.strip() or "Team"
    closing_name = sender_name.strip() or "Your Name"

    summary_block = normalize_value(meeting.get("summary"), "No summary generated.")
    objective = normalize_value(meeting.get("objective"), "Not provided")
    outcome = normalize_value(meeting.get("outcome"), "Not provided")
    follow_up_line = f"Follow-up: {'Yes' if meeting.get('followUp') else 'No'}"

    actions = meeting.get("actions", [])
    action_lines = []
    if actions:
        for action in actions:
            action_lines.append(
                f"- {normalize_value(action.get('text'))} | Assignee: {normalize_value(action.get('owner'), 'Not stated')} | "
                f"Deadline: {normalize_value(action.get('deadline'), 'None')} | Status: {normalize_status(action)}"
            )

    decisions = meeting.get("keyDecisions", [])
    decision_lines = [f"- {normalize_value(item)}" for item in decisions] if decisions else []

    lines = [
        f"Dear {greeting_name},",
        "",
        "Please find below the meeting summary report for your reference.",
        "",
        summary_block,
        "",
        f"Objective: {objective}",
        f"Outcome: {outcome}",
        follow_up_line,
    ]

    if decision_lines:
        lines.extend(["", "Key Decisions:"])
        lines.extend(decision_lines)

    if action_lines:
        lines.extend(["", "Action Items:"])
        lines.extend(action_lines)

    lines.extend(["", "Regards,", closing_name])
    return "\n".join(lines)


# ============================================================
# Section 4. UI Components
# ============================================================

def render_email_copy_block(meeting: dict, key_prefix: str) -> None:
    st.markdown(
        """
        <div class="section-card">
            <div class="mini-title">Email Copy Template</div>
            <div class="mini-copy">Fill in the names below, then copy the prepared subject and email body into Outlook or any email app.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    recipient_name = st.text_input(
        "Recipient Name",
        key=f"{key_prefix}_recipient_name",
        placeholder="Recipient name",
    )
    sender_name = st.text_input(
        "Sender Name",
        key=f"{key_prefix}_sender_name",
        placeholder="Your name",
    )
    st.text_input(
        "Email Subject",
        value=build_meeting_email_subject(meeting),
        key=f"{key_prefix}_subject",
    )
    st.text_area(
        "Email Body",
        value=build_meeting_email_body(meeting, recipient_name=recipient_name, sender_name=sender_name),
        height=260,
        key=f"{key_prefix}_body",
    )

def render_dashboard_chat(meetings: list) -> None:
    st.markdown('<div class="chat-thread dashboard-chat-thread">', unsafe_allow_html=True)
    if st.session_state.chat_history:
        for message in st.session_state.chat_history[-4:]:
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


# ============================================================
# Section 5. Meeting Analysis Prompt
# ============================================================
PIPELINE_SYSTEM = """You are a meeting intelligence system. Return ONLY valid JSON.

Goals:
- Extract persons, organizations, dates, and locations.
- Identify the meeting objective from the transcript and metadata.
- Write a concise but complete 4-5 sentence summary.
- Extract key decisions, discussion points, and action items.
- Mark follow-up as true only when something is still pending.

Rules:
- Treat only explicitly stated tasks, requests, assignments, and pending items as action items.
- Do not infer hidden or implied tasks from general discussion or meeting purpose.
- Only keep action items that belong to TalentCorp or TalentCorp internal teams. If the assignee or responsibility is clearly for another company, do not place it in action_items.
- External-party responsibilities can still appear in the summary or discussion points, but not as TalentCorp action items.
- Always return at least 1-3 discussion_points when the transcript contains actual meeting content.
- discussion_points should capture the main topics discussed, reviewed, aligned, explored, or presented.
- Return key_decisions only when the transcript clearly states a decision, agreement, confirmation, or approval.
- If a recap includes important dates, months, launch periods, deadlines, or timelines, include them only when they are explicitly mentioned.
- If owner or deadline is missing, use "Not stated" and "None".
- Prefer separate action items instead of merging unrelated tasks, but only when each task is explicitly stated.
- If structured metadata is provided, use it as context.
- If the recap only describes the purpose of a meeting, expected outcome, or general discussion without a direct task, return an empty action_items list and set follow_up to false unless a pending task is clearly stated.

Return this schema only:
{
  "title": "string",
  "meeting_type": "string",
  "category": "string",
  "nlp_pipeline": {
    "token_count": 0,
    "sentence_count": 0,
    "named_entities": {
      "persons": [],
      "organizations": [],
      "dates": [],
      "locations": []
    }
  },
  "classification": {
    "action_items_count": 0,
    "decisions_count": 0,
    "discussion_points_count": 0
  },
  "objective": "string",
  "summary": "string",
  "outcome": "string",
  "follow_up": true,
  "follow_up_reason": "string",
  "key_decisions": [],
  "discussion_points": [],
  "action_items": [
    {
      "text": "string",
      "owner": "Not stated",
      "company": "string",
      "deadline": "None",
      "priority": "High|Medium|Low",
      "follow_up_required": true,
      "follow_up_reason": "string",
      "suggestion": "string",
      "ner_entities": []
    }
  ],
  "estimated_budget": 0,
  "budget_notes": ""
}
"""


def extract_json(raw: str) -> dict:
    cleaned = re.sub(r"```(?:json)?", "", raw).strip()
    cleaned = cleaned.replace("\u201c", '"').replace("\u201d", '"').replace("\u2018", "'").replace("\u2019", "'")

    def try_load(candidate: str) -> dict:
        return json.loads(candidate)

    def try_literal(candidate: str) -> dict:
        parsed = ast.literal_eval(candidate)
        if isinstance(parsed, dict):
            return parsed
        raise ValueError("Recovered content is not a dictionary")

    try:
        return try_load(cleaned)
    except json.JSONDecodeError:
        match = re.search(r"\{.*\}", cleaned, re.DOTALL)
        if match:
            candidate = match.group(0).strip()
            quote_bare_values = lambda text: re.sub(
                r"([:\[,]\s*)([A-Za-z_][A-Za-z0-9_\-/ ]*)(?=\s*[,}\]])",
                lambda m: f"{m.group(1)}{json.dumps(m.group(2).strip())}",
                text,
            )
            quote_bare_keys = lambda text: re.sub(
                r'([{\s,])([A-Za-z_][A-Za-z0-9_\- ]*)(\s*:)',
                lambda m: f'{m.group(1)}"{m.group(2).strip()}"{m.group(3)}',
                text,
            )
            repairs = [
                candidate,
                re.sub(r",\s*([}\]])", r"\1", candidate),
                quote_bare_values(candidate),
                re.sub(r",\s*([}\]])", r"\1", quote_bare_values(candidate)),
                quote_bare_keys(candidate),
                re.sub(r",\s*([}\]])", r"\1", quote_bare_keys(candidate)),
                quote_bare_values(quote_bare_keys(candidate)),
                re.sub(r",\s*([}\]])", r"\1", quote_bare_values(quote_bare_keys(candidate))),
            ]
            for repaired in repairs:
                try:
                    return try_load(repaired)
                except json.JSONDecodeError:
                    try:
                        return try_literal(repaired)
                    except Exception:
                        continue
        raise


def recover_json_with_ollama(raw: str) -> dict:
    repair_system = (
        "You repair malformed meeting-analysis JSON. "
        "Return only valid JSON with the same meaning. No markdown, no explanation."
    )
    repair_prompt = f"Fix this into valid JSON only:\n\n{raw[:2500]}"
    repaired = call_ollama(repair_system, repair_prompt, max_tokens=150)
    return extract_json(repaired)


def build_safe_pipeline_result(transcript: str, metadata: dict | None = None) -> dict:
    metadata = metadata or {}
    discussion_points = fallback_discussion_points(transcript)
    key_decisions = fallback_key_decisions(transcript)
    title = normalize_value(metadata.get("Title"), "") or normalize_value(metadata.get("Activity Title"), "") or "Untitled"
    meeting_type = normalize_value(metadata.get("Activity Type"), "") or "Not Provided"
    category = normalize_value(metadata.get("Category"), "") or "Not Provided"
    objective = discussion_points[0] if discussion_points else "Objective not clearly extracted."
    summary_sentences = transcript_sentences(transcript)[:3]
    summary = " ".join(summary_sentences).strip() or "Summary could not be generated from the transcript."
    return {
        "title": title,
        "meeting_type": meeting_type,
        "category": category,
        "nlp_pipeline": {
            "token_count": 0,
            "sentence_count": len(transcript_sentences(transcript)),
            "named_entities": {
                "persons": [],
                "organizations": [],
                "dates": [],
                "locations": [],
            },
        },
        "classification": {
            "action_items_count": 0,
            "decisions_count": len(key_decisions),
            "discussion_points_count": len(discussion_points),
        },
        "objective": objective,
        "summary": summary,
        "outcome": "Not provided",
        "follow_up": False,
        "follow_up_reason": "",
        "key_decisions": key_decisions,
        "discussion_points": discussion_points,
        "action_items": [],
        "estimated_budget": 0,
        "budget_notes": "",
    }


def normalize_pipeline_result(result: dict, transcript: str, metadata: dict | None = None) -> dict:
    safe = build_safe_pipeline_result(transcript, metadata)
    if not isinstance(result, dict):
        return safe
    merged = safe | result
    merged["nlp_pipeline"] = {
        **safe["nlp_pipeline"],
        **(result.get("nlp_pipeline", {}) if isinstance(result.get("nlp_pipeline", {}), dict) else {}),
    }
    safe_entities = safe["nlp_pipeline"]["named_entities"]
    result_entities = merged["nlp_pipeline"].get("named_entities", {})
    merged["nlp_pipeline"]["named_entities"] = {
        **safe_entities,
        **(result_entities if isinstance(result_entities, dict) else {}),
    }
    merged["classification"] = {
        **safe["classification"],
        **(result.get("classification", {}) if isinstance(result.get("classification", {}), dict) else {}),
    }
    for key in ("key_decisions", "discussion_points", "action_items"):
        if not isinstance(merged.get(key), list):
            merged[key] = safe[key]
    for key in ("title", "meeting_type", "category", "objective", "summary", "outcome", "budget_notes"):
        merged[key] = normalize_value(merged.get(key), safe[key])
    merged["estimated_budget"] = merged.get("estimated_budget", 0) or 0
    merged["follow_up"] = parse_yes_no(merged.get("follow_up"))
    merged["follow_up_reason"] = normalize_value(merged.get("follow_up_reason"), "")
    return merged


def run_pipeline(transcript: str, metadata: dict | None = None) -> dict:
    cleaned_transcript = compact_transcript_for_prompt(transcript.strip(), max_chars=800)
    objective_only = is_objective_only_transcript(cleaned_transcript)
    metadata_lines = []
    for label, value in (metadata or {}).items():
        normalized = normalize_value(value, "")
        if normalized:
            metadata_lines.append(f"{label}: {normalized}")
    metadata_block = "\n".join(metadata_lines)
    objective_note = (
        "This transcript is objective-only. Return no action items and set follow_up to false.\n"
        if objective_only
        else ""
    )
    user_msg = (
        "Return concise JSON with objective, summary, follow-up, and action items.\n"
        "Use only explicit tasks from the text. Do not invent implied action items.\n"
        f"{objective_note}"
        f"Activity metadata:\n{metadata_block or 'None provided'}\n\n"
        f"Meeting content:\n{cleaned_transcript}"
    )
    raw = call_ollama(PIPELINE_SYSTEM, user_msg, max_tokens=250)
    try:
        result = extract_json(raw)
    except Exception:
        try:
            result = recover_json_with_ollama(raw)
        except Exception:
            result = build_safe_pipeline_result(cleaned_transcript, metadata)
    result = normalize_pipeline_result(result, cleaned_transcript, metadata)
    action_count = len(result.get("action_items", []))
    if not objective_only and action_count > 0:
        result["follow_up"] = True
        if not str(result.get("follow_up_reason", "")).strip():
            result["follow_up_reason"] = "Action items are still pending."
    if objective_only:
        result["action_items"] = []
        result["follow_up"] = False
        result["follow_up_reason"] = ""
        result.setdefault("classification", {})
        result["classification"]["action_items_count"] = 0
        result["classification"]["decisions_count"] = len(result.get("key_decisions", []))
        result["classification"]["discussion_points_count"] = len(result.get("discussion_points", []))

    if not result.get("discussion_points"):
        result["discussion_points"] = fallback_discussion_points(cleaned_transcript)
    if not result.get("key_decisions"):
        result["key_decisions"] = fallback_key_decisions(cleaned_transcript)

    result["follow_up_reason"] = ""
    filtered_actions = filter_talentcorp_actions(result.get("action_items", []))
    if not filtered_actions and not objective_only:
        filtered_actions = fallback_action_items(cleaned_transcript)
    result["action_items"] = filtered_actions
    result.setdefault("classification", {})
    result["classification"]["action_items_count"] = len(filtered_actions)
    result["classification"]["decisions_count"] = len(result.get("key_decisions", []))
    result["classification"]["discussion_points_count"] = len(result.get("discussion_points", []))
    if not filtered_actions:
        result["follow_up"] = False
    return result


def chat_with_meetings(question: str, meetings: list) -> str:
    question_lower = question.lower().strip()
    question_tokens = {
        token
        for token in re.findall(r"[a-zA-Z0-9&]+", question_lower)
        if len(token) >= 3 and token not in {
            "what",
            "when",
            "where",
            "which",
            "there",
            "their",
            "about",
            "with",
            "from",
            "have",
            "that",
            "this",
            "item",
            "items",
            "task",
            "tasks",
            "action",
            "actions",
            "meeting",
            "meetings",
        }
    }

    def meeting_search_blob(meeting: dict) -> str:
        return " ".join(
            [
                normalize_value(meeting.get("title"), ""),
                normalize_value(meeting.get("summary"), ""),
                normalize_value(meeting.get("outcome"), ""),
                normalize_value(meeting.get("meetingID"), ""),
                normalize_value(meeting.get("activityId"), ""),
                join_list(meeting.get("stakeholders", []), ""),
                join_list(meeting.get("companies", []), ""),
                join_list(meeting.get("discussionPoints", []), ""),
                join_list(meeting.get("keyDecisions", []), ""),
            ]
        ).lower()

    scored_meetings = []
    for meeting in meetings:
        blob = meeting_search_blob(meeting)
        score = sum(1 for token in question_tokens if token in blob)
        if score > 0:
            scored_meetings.append((score, meeting))

    relevant_meetings = [meeting for _, meeting in sorted(scored_meetings, key=lambda item: item[0], reverse=True)]
    if not relevant_meetings:
        relevant_meetings = meetings[:5]

    action_question = any(
        keyword in question_lower
        for keyword in ["action", "task", "deadline", "owner", "pending", "follow up", "follow-up"]
    )

    if action_question and relevant_meetings:
        top_meeting = relevant_meetings[0]
        top_title = normalize_value(top_meeting.get("title"), "this meeting")
        active_actions = top_meeting.get("actions", [])
        if active_actions:
            lines = [
                f"For the meeting \"{top_title}\", the action items are:"
            ]
            for action in active_actions:
                lines.append(
                    f"- {normalize_value(action.get('text'))} | owner: {normalize_value(action.get('owner'), 'Not stated')} | "
                    f"status: {normalize_status(action)} | deadline: {normalize_value(action.get('deadline'), 'None')}"
                )
            return "\n".join(lines)
        return f'There is no action item mentioned in the meeting data for "{top_title}".'

    meeting_blocks = []
    for meeting in relevant_meetings[:5]:
        actions = meeting.get("actions", [])
        action_lines = [
            f"- {normalize_value(action.get('text'))} | owner: {normalize_value(action.get('owner'), 'Not stated')} | "
            f"status: {normalize_status(action)} | deadline: {normalize_value(action.get('deadline'), 'None')}"
            for action in actions
        ]
        meeting_blocks.append(
            "\n".join(
                [
                    f"Date: {meeting['date']}",
                    f"Title: {meeting['title']}",
                    f"Outcome: {meeting.get('outcome', '')}",
                    f"Follow-up: {meeting.get('followUp', False)}",
                    "Action items:",
                    "\n".join(action_lines) if action_lines else "- None",
                ]
            )
        )

    ctx = "\n\n".join(meeting_blocks) if meeting_blocks else "No meeting data available."
    system = """You are MeetIQ's AI assistant.

You have two responsibilities:
1. Answer questions using stored meeting data whenever the question is about meetings, tasks, follow-up items, decisions, owners, or history.
2. If the user asks for general help, guidance, suggestions, templates, or advice that is not directly available in the meeting data, answer helpfully using your general knowledge.

RULES:
- If the user asks about pending, open, overdue, incomplete, unresolved, or outstanding actions, treat statuses Pending, In Progress, and Overdue as not completed.
- Never say there are no pending items if meeting data contains Pending, In Progress, or Overdue actions.
- When answering from meeting data, mention the relevant meeting title, owner, deadline, and status when available.
- If the question is broader than the stored data, first answer from the data if relevant, then provide practical suggestions.
- Be concise, practical, and business-friendly.
"""
    user_msg = f"Meeting data:\n{ctx}\n\nQuestion: {question}"
    return call_ollama(system, user_msg, max_tokens=260)


# ============================================================
# Section 6. App Theme and State
# ============================================================

st.set_page_config(page_title="MeetIQ", layout="wide")

st.markdown(
    """
    <style>
    :root {
        --bg-outer: #2d3342;
        --surface: #ffffff;
        --surface-soft: #f1f5f9;
        --border: #cbd5e1;
        --text: #0f172a;
        --text-muted: #334155;
        --text-soft: #64748b;
        --brand: #123a63;
        --brand-2: #4f46e5;
        --accent: #0f766e;
    }
    .stApp {
        background:
            radial-gradient(circle at top right, rgba(79, 70, 229, 0.12), transparent 30%),
            radial-gradient(circle at bottom left, rgba(15, 118, 110, 0.10), transparent 35%),
            var(--bg-outer);
        color: var(--text);
        font-family: "Aptos", "Segoe UI", Arial, sans-serif;
        font-size: 16px;
        line-height: 1.6;
    }
    .block-container {
        padding: 1.25rem 1.5rem 2rem;
        margin: 1rem auto;
        max-width: 1380px;
        background: linear-gradient(180deg, var(--surface) 0%, #fbfdff 100%);
        border-radius: 28px;
        box-shadow: 0 22px 60px rgba(15, 23, 42, 0.18);
    }
    .hero-shell {
        background: linear-gradient(135deg, #0f172a 0%, #1e3a5f 42%, #4f46e5 100%);
        color: white;
        border-radius: 24px;
        padding: 1.45rem 1.8rem 1.35rem;
        margin-bottom: 1.15rem;
        box-shadow: 0 18px 40px rgba(15, 23, 42, 0.18);
        position: relative;
        overflow: hidden;
    }
    .hero-shell::before {
        content: "";
        position: absolute;
        top: -30px;
        right: 22%;
        width: 120px;
        height: 120px;
        border-radius: 50%;
        background: rgba(255, 255, 255, 0.10);
        filter: blur(2px);
    }
    .hero-shell::after {
        content: "";
        position: absolute;
        inset: auto -40px -40px auto;
        width: 180px;
        height: 180px;
        border-radius: 50%;
        background: rgba(255, 255, 255, 0.08);
    }
    .hero-shell h1 {
        margin: 0;
        font-size: 2.15rem;
        line-height: 1.1;
        letter-spacing: 0.02em;
        color: #ffffff !important;
        text-align: left;
        font-weight: 800;
        font-family: "Trebuchet MS", "Segoe UI", "Verdana", sans-serif;
        position: relative;
        z-index: 1;
        max-width: 54rem;
    }
    .hero-shell p {
        margin: 0.35rem 0 0;
        color: rgba(255,255,255,0.92);
        max-width: 42rem;
        text-align: left;
        font-size: 1rem;
        position: relative;
        z-index: 1;
        font-weight: 700;
    }
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #4657c8 0%, #3246af 45%, #24338d 100%);
        border-right: 1px solid rgba(255,255,255,0.10);
    }
    section[data-testid="stSidebar"] .block-container {
        padding-top: 1rem;
        padding-bottom: 1rem;
    }
    .sidebar-title {
        margin: 0 0 1rem;
        color: #ffffff !important;
        font-size: 1.45rem;
        font-weight: 800;
        letter-spacing: 0.01em;
    }
    .sidebar-subtitle {
        margin: -0.7rem 0 1rem;
        color: rgba(255,255,255,0.78);
        font-size: 0.88rem;
    }
    section[data-testid="stSidebar"] .stButton > button {
        border-radius: 18px !important;
        min-height: 3rem;
        border: 0 !important;
        font-weight: 700 !important;
        box-shadow: none !important;
        background: rgba(255,255,255,0.12) !important;
        color: #ffffff !important;
        backdrop-filter: blur(4px);
    }
    [data-testid="stFileUploader"] button,
    [data-testid="stFileUploader"] button span,
    [data-testid="stFileUploader"] section button,
    [data-testid="stFileUploaderDropzone"] button,
    [data-testid="stFileUploaderDropzone"] button span {
        color: #ffffff !important;
    }
    .card-link {
        text-decoration: none !important;
        display: block;
        color: inherit !important;
    }
    .clickable-card {
        cursor: pointer;
        transition: transform 0.15s ease, box-shadow 0.15s ease, border-color 0.15s ease;
    }
    .clickable-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 14px 28px rgba(15, 23, 42, 0.09);
        border-color: #b8c8dd;
    }
    h2, h3, label, .stMarkdown, .stCaption, .stRadio label, .stSelectbox label {
        color: var(--text) !important;
    }
    h2, h3 {
        letter-spacing: -0.02em;
    }
    .stTextArea label, .stTextInput label, .stNumberInput label, .stFileUploader label {
        color: var(--text) !important;
        font-weight: 600 !important;
    }
    .stRadio div[role="radiogroup"] label, .stCheckbox label {
        color: var(--text-muted) !important;
    }
    .stRadio div[role="radiogroup"] p {
        color: var(--text) !important;
    }
    .st-emotion-cache-16idsys p, .st-emotion-cache-1r6slb0, .stAlert {
        color: var(--text) !important;
    }
    .hero-panel, .kpi-card, .action-card, .info-card {
        background: var(--surface);
        border: 1px solid var(--border);
        border-radius: 18px;
        box-shadow: 0 14px 28px rgba(15, 23, 42, 0.06);
    }
    .hero-panel {
        padding: 1.35rem 1.4rem;
    }
    .hero-badge {
        display: inline-block;
        background: #e2e8f0;
        color: var(--brand);
        padding: 0.35rem 0.7rem;
        border-radius: 999px;
        font-size: 0.8rem;
        font-weight: 700;
        margin-bottom: 0.8rem;
    }
    .hero-panel h2 {
        margin: 0 0 0.6rem;
        color: var(--text);
        font-size: 1.55rem;
    }
    .hero-panel p {
        margin: 0;
        color: var(--text-soft);
        line-height: 1.6;
    }
    .hero-grid {
        display: grid;
        grid-template-columns: repeat(2, minmax(0, 1fr));
        gap: 0.9rem;
        margin-top: 1.1rem;
        color: var(--text-muted);
    }
    .kpi-card {
        padding: 1rem 1.05rem;
        min-height: 108px;
    }
    .nav-card-marker + div[data-testid="stButton"] > button {
        min-height: 168px !important;
        border-radius: 22px !important;
        border: 1px solid #d7e3f3 !important;
        background: #ffffff !important;
        color: #0f172a !important;
        box-shadow: 0 16px 28px rgba(30, 58, 95, 0.08) !important;
        white-space: pre-line !important;
        text-align: left !important;
        justify-content: flex-start !important;
        align-items: flex-start !important;
        padding: 1.1rem 1.1rem !important;
        font-weight: 700 !important;
        line-height: 1.55 !important;
    }
    .nav-card-marker + div[data-testid="stButton"] > button:hover {
        border-color: #9db6db !important;
        box-shadow: 0 18px 32px rgba(79, 70, 229, 0.12) !important;
        color: #0f172a !important;
    }
    .completion-marker + div[data-testid="stButton"] > button {
        min-height: 200px !important;
    }
    .nav-card-marker + div[data-testid="stButton"] > button p {
        color: #0f172a !important;
        font-size: 1rem !important;
        line-height: 1.6 !important;
        margin: 0 !important;
    }
    .kpi-label {
        color: var(--text-soft);
        font-size: 0.86rem;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        font-weight: 700;
    }
    .kpi-value {
        margin-top: 0.45rem;
        font-size: 1.85rem;
        line-height: 1;
        font-weight: 800;
        color: var(--text);
    }
    .kpi-subtitle {
        margin-top: 0.45rem;
        color: var(--text-soft);
        font-size: 0.92rem;
    }
    .action-card {
        padding: 1rem 1rem 0.85rem;
        margin-bottom: 0.75rem;
    }
    .action-top {
        display: flex;
        align-items: start;
        justify-content: space-between;
        gap: 0.8rem;
        margin-bottom: 0.45rem;
    }
    .action-title {
        color: var(--text);
        font-weight: 700;
        font-size: 1rem;
    }
    .action-meta {
        color: var(--text-muted);
        font-size: 0.92rem;
        margin-bottom: 0.28rem;
    }
    .action-subtle {
        color: var(--text-soft);
        font-size: 0.9rem;
    }
    .info-card {
        padding: 1rem;
        margin-bottom: 0.9rem;
    }
    .section-card {
        padding: 1rem;
        margin-top: 0.9rem;
        margin-bottom: 0.9rem;
        background: var(--surface);
        border: 1px solid var(--border);
        border-radius: 18px;
        box-shadow: 0 14px 28px rgba(15, 23, 42, 0.06);
    }
    .mini-title {
        color: var(--text);
        font-size: 1rem;
        font-weight: 700;
        margin-bottom: 0.35rem;
    }
    .mini-copy {
        color: var(--text-soft);
        font-size: 0.96rem;
        line-height: 1.55;
    }
    .stButton button {
        background: var(--brand);
        color: #ffffff;
        border: none;
        border-radius: 12px;
        padding: 0.55rem 0.9rem;
        box-shadow: none;
        transition: none;
    }
    .stTextInput input, .stNumberInput input, .stTextArea textarea, .stSelectbox div[data-baseweb="select"] {
        border-radius: 12px !important;
    }
    .stTextInput input, .stNumberInput input, .stTextArea textarea {
        color: #ffffff !important;
        caret-color: #ffffff !important;
    }
    .stSelectbox div[data-baseweb="select"] * {
        color: #ffffff !important;
    }
    .stDateInput input, .stDateInput * {
        color: #ffffff !important;
    }
    .stTextInput input::placeholder, .stNumberInput input::placeholder, .stTextArea textarea::placeholder {
        color: rgba(255, 255, 255, 0.72) !important;
    }
    .stChatInput textarea, div[data-testid="stChatInput"] textarea {
        color: #ffffff !important;
        caret-color: #ffffff !important;
    }
    .stChatInput textarea::placeholder, div[data-testid="stChatInput"] textarea::placeholder {
        color: rgba(255, 255, 255, 0.72) !important;
    }
    .chat-thread {
        display: flex;
        flex-direction: column;
        gap: 0.85rem;
        margin: 0.9rem 0 1rem;
    }
    .chat-bubble {
        max-width: min(82%, 820px);
        padding: 0.9rem 1rem;
        border-radius: 18px;
        box-shadow: 0 10px 24px rgba(15, 23, 42, 0.06);
        border: 1px solid var(--border);
        line-height: 1.55;
        font-size: 0.98rem;
        word-wrap: break-word;
        white-space: pre-wrap;
    }
    .chat-bubble.user {
        margin-left: auto;
        background: linear-gradient(135deg, #3b82f6, #2563eb);
        color: #ffffff;
        border-color: rgba(37, 99, 235, 0.22);
        border-bottom-right-radius: 6px;
    }
    .chat-bubble.assistant {
        margin-right: auto;
        background: #f3f4f6;
        color: var(--text);
        border-bottom-left-radius: 6px;
    }
    .chat-bubble.user strong, .chat-bubble.user b {
        color: #ffffff;
    }
    .streamlit-expanderContent p, .streamlit-expanderContent div, .streamlit-expanderContent label, .streamlit-expanderContent span {
        color: var(--text) !important;
    }
    [data-testid="stExpander"] summary, [data-testid="stExpander"] summary p, [data-testid="stExpander"] summary span, [data-testid="stExpander"] summary div {
        color: var(--text) !important;
        opacity: 1 !important;
    }
    .date-group {
        margin: 1rem 0 0.75rem;
        padding: 0.65rem 0.9rem;
        border-radius: 12px;
        background: #dbe7f3;
        color: var(--brand);
        font-weight: 700;
        border: 1px solid #c6d4e3;
    }
    .dashboard-shell {
        display: grid;
        grid-template-columns: minmax(0, 1.35fr) minmax(310px, 0.85fr);
        gap: 1rem;
        align-items: start;
    }
    .dashboard-stack {
        display: flex;
        flex-direction: column;
        gap: 1rem;
    }
    .dashboard-card {
        background: rgba(255,255,255,0.94);
        border: 1px solid var(--border);
        border-radius: 22px;
        box-shadow: 0 16px 32px rgba(15, 23, 42, 0.08);
        padding: 1rem 1.05rem;
    }
    .dashboard-title {
        font-size: 1.55rem;
        font-weight: 800;
        color: var(--text);
        margin: 0 0 0.2rem;
    }
    .dashboard-copy {
        color: var(--text-soft);
        font-size: 0.96rem;
        margin: 0;
    }
    .search-shell {
        display: flex;
        align-items: center;
        gap: 0.8rem;
        background: #f8fbff;
        border: 1px solid #d8e2f1;
        border-radius: 18px;
        padding: 0.55rem 0.85rem;
        margin-bottom: 1rem;
    }
    .search-icon {
        width: 2.6rem;
        height: 2.6rem;
        border-radius: 50%;
        display: grid;
        place-items: center;
        background: linear-gradient(135deg, var(--brand), var(--brand-2));
        color: #fff;
        font-size: 1.1rem;
        flex: 0 0 auto;
    }
    .calendar-widget {
        margin-top: 0.6rem;
    }
    .calendar-grid {
        display: grid;
        grid-template-columns: repeat(7, minmax(0, 1fr));
        gap: 0.42rem;
    }
    .calendar-head {
        margin-bottom: 0.45rem;
    }
    .calendar-day-label {
        text-align: center;
        color: var(--text-soft);
        font-size: 0.78rem;
        font-weight: 700;
        text-transform: uppercase;
    }
    .calendar-day {
        aspect-ratio: 1 / 1;
        border-radius: 16px;
        display: grid;
        place-items: center;
        background: #f3f6fb;
        color: var(--text);
        font-weight: 700;
        border: 1px solid #e2e8f0;
    }
    .calendar-day.empty {
        background: transparent;
        border-color: transparent;
    }
    .calendar-day.today {
        background: #1e3a5f;
        color: #ffffff;
    }
    .calendar-day.pending-deadline {
        background: #fef3c7;
        border-color: #f59e0b;
        color: #92400e;
        box-shadow: inset 0 0 0 2px #facc15;
    }
    .calendar-day.pending-deadline.today {
        background: linear-gradient(135deg, #fef3c7 0%, #fde68a 100%);
        color: #92400e;
        border-color: #f59e0b;
    }
    .upcoming-item {
        padding: 0.85rem 0.95rem;
        border-radius: 16px;
        background: #f8fbff;
        border: 1px solid #d8e2f1;
        margin-top: 0.7rem;
    }
    .upcoming-top {
        display: flex;
        justify-content: space-between;
        gap: 0.75rem;
        align-items: start;
    }
    .upcoming-date {
        min-width: 74px;
        padding: 0.45rem 0.55rem;
        border-radius: 14px;
        background: linear-gradient(135deg, #e0e7ff, #dbeafe);
        color: #1e3a5f;
        text-align: center;
        font-weight: 800;
        font-size: 0.82rem;
    }
    .search-result-card {
        padding: 0.95rem 1rem;
        border-radius: 18px;
        background: #ffffff;
        border: 1px solid #dbe4f0;
        box-shadow: 0 10px 22px rgba(15, 23, 42, 0.05);
        margin-bottom: 0.85rem;
    }
    .search-result-top {
        display: flex;
        justify-content: space-between;
        align-items: start;
        gap: 0.85rem;
        margin-bottom: 0.45rem;
    }
    .result-pill {
        white-space: nowrap;
        padding: 0.3rem 0.7rem;
        border-radius: 999px;
        background: #eef2ff;
        color: #3730a3;
        font-size: 0.8rem;
        font-weight: 700;
    }
    .dashboard-chat-thread .chat-bubble {
        max-width: 100%;
        font-size: 0.92rem;
    }
    .section-link {
        color: #4f46e5;
        font-size: 0.9rem;
        font-weight: 700;
        text-decoration: none;
    }
    .completion-card {
        background: #ffffff;
        border: 1px solid #d9e2ec;
        border-radius: 16px;
        box-shadow: 0 10px 24px rgba(15, 23, 42, 0.05);
        padding: 1rem 1.05rem;
        min-height: 108px;
    }
    .completion-wrap {
        display: flex;
        justify-content: center;
        padding: 0.2rem 0 0.35rem;
    }
    .completion-ring {
        --size: 88px;
        width: var(--size);
        height: var(--size);
        border-radius: 50%;
        background: conic-gradient(#d9ff57 calc(var(--pct) * 1%), #eef2f7 0);
        display: grid;
        place-items: center;
        box-shadow: inset 0 0 0 1px rgba(15, 23, 42, 0.05);
    }
    .completion-inner {
        width: 60px;
        height: 60px;
        border-radius: 50%;
        background: #ffffff;
        display: grid;
        place-items: center;
        font-size: 1.2rem;
        font-weight: 800;
        color: #4f46e5;
    }
    @media (max-width: 980px) {
        .dashboard-shell {
            grid-template-columns: 1fr;
        }
        .app-shell {
            grid-template-columns: 1fr;
        }
    }
    div[data-baseweb="input"], div[data-baseweb="select"], textarea, input {
        color: var(--text) !important;
        font-size: 1rem !important;
    }
    div[data-baseweb="input"]:focus-within, div[data-baseweb="select"]:focus-within, textarea:focus, input:focus {
        outline: 3px solid rgba(29, 78, 216, 0.22) !important;
        outline-offset: 2px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def init_state():
    if "data_loaded" not in st.session_state:
        meetings, departments = load_app_data()
        st.session_state.meetings = meetings
        st.session_state.departments = departments
        st.session_state.data_loaded = True
    if "meetings" not in st.session_state:
        st.session_state.meetings = []
    if "departments" not in st.session_state:
        st.session_state.departments = []
    if "chat_history" not in st.session_state:
        st.session_state.chat_history = []
    if "pending_result" not in st.session_state:
        st.session_state.pending_result = None
    if "capture_transcript" not in st.session_state:
        st.session_state.capture_transcript = ""
    if "capture_activity_id" not in st.session_state:
        st.session_state.capture_activity_id = ""
    if "current_page" not in st.session_state:
        st.session_state.current_page = "Dashboard"
    if "tracker_focus" not in st.session_state:
        st.session_state.tracker_focus = "all"


init_state()
seed_default_departments()

meetings = st.session_state.meetings
meeting_df = build_meeting_dataframe(meetings)
action_df = build_action_dataframe(meetings)

with st.sidebar:
    st.markdown(
        """
        <div class="sidebar-title">AI-Powered Meeting Insight Generator &amp; Action Tracker</div>
        <div class="sidebar-subtitle">for Talentcorp by Zaf &lt;3</div>
        """,
        unsafe_allow_html=True,
    )
    st.button(
        "Dashboard",
        key="nav_dashboard",
        use_container_width=True,
        type="primary" if st.session_state.current_page == "Dashboard" else "secondary",
        on_click=set_current_page,
        args=("Dashboard",),
    )
    st.button(
        "Capture",
        key="nav_capture",
        use_container_width=True,
        type="primary" if st.session_state.current_page == "Capture" else "secondary",
        on_click=set_current_page,
        args=("Capture",),
    )
    st.button(
        "Tracker",
        key="nav_tracker",
        use_container_width=True,
        type="primary" if st.session_state.current_page == "Tracker" else "secondary",
        on_click=set_current_page,
        args=("Tracker",),
    )
    st.button(
        "Finance",
        key="nav_finance",
        use_container_width=True,
        type="primary" if st.session_state.current_page == "Finance" else "secondary",
        on_click=set_current_page,
        args=("Finance",),
    )


# ============================================================
# Section 7. Capture Tab
# ============================================================

if st.session_state.current_page == "Capture":
    st.subheader("Capture & Analyze")
    st.caption("Paste notes, upload audio, or record a meeting to generate a structured executive brief.")

    activity_box = st.container(border=True)
    with activity_box:
        act_left, act_right = st.columns(2)
        dept_names = get_department_options()
        with act_left:
            activity_category = st.selectbox("Category", CATEGORIES, key="capture_activity_category")
            activity_id = st.text_input("Activity ID", key="capture_activity_id", placeholder="Generate or enter activity ID")
            st.button("Generate Activity ID", key="generate_activity_id_btn", on_click=set_generated_activity_id)
            st.button("Clear Activity ID", key="clear_activity_id_btn", on_click=clear_generated_activity_id)
            activity_title = st.text_input("Title", key="capture_activity_title", placeholder="Enter meeting or activity title")
            activity_type = st.selectbox("Activity Type", ACTIVITY_TYPE_OPTIONS, key="capture_activity_type")
        with act_right:
            meeting_date = st.date_input("Meeting Date", value=date.today(), key="capture_meeting_date")
            actual_cost = st.number_input("Actual Cost (RM)", min_value=0.0, step=50.0)
            dept_choice = st.selectbox("Department", dept_names)

    role = ""
    main_activity = ""
    link_photo = ""
    link_photo_url = ""
    organization_type = ""
    date_from = meeting_date
    date_to = meeting_date
    district = ""
    invitation_from = ""
    location_meeting = ""
    other_reps = ""
    representative_position = ""
    representative_name = ""
    representative_department = ""
    stfemail = ""
    supemail = ""
    updated_by = ""

    transcript_box = st.container(border=True)
    with transcript_box:
        audio_mode = st.radio(
            "Audio source",
            ["Manual transcript", "Upload audio file", "Record meeting audio"],
            horizontal=True,
        )
        transcript_mode = st.selectbox(
            "Transcript output",
            ["Translate to English", "Keep spoken language"],
            help="Use local Whisper translation for mixed-language meetings before summarization.",
        )
        document_files = st.file_uploader(
            "Add supporting files",
            type=["pdf", "docx", "xlsx", "xls", "csv"],
            help="Upload multiple PDF, Word, Excel, or CSV files to append into the meeting transcript.",
            accept_multiple_files=True,
        )

        uploaded_audio = None
        recorded_audio = None
        if audio_mode == "Upload audio file":
            uploaded_audio = st.file_uploader(
                "Upload audio",
                type=["mp3", "m4a", "wav", "mp4", "mpeg", "mpga", "webm"],
                help="Supported for local Whisper transcription.",
            )
        elif audio_mode == "Record meeting audio":
            if hasattr(st, "audio_input"):
                recorded_audio = st.audio_input("Record meeting")
            else:
                st.info("This Streamlit version does not support in-app recording yet. Use audio upload instead.")

        transcribe_clicked = st.button(
            "Transcribe Audio",
            disabled=audio_mode == "Manual transcript" or (uploaded_audio is None and recorded_audio is None),
        )

        if transcribe_clicked:
            audio_source = uploaded_audio if uploaded_audio is not None else recorded_audio
            with st.spinner("Transcribing audio with local Whisper..."):
                try:
                    st.session_state.capture_transcript = transcribe_audio_file(
                        audio_source,
                        translate_to_english=transcript_mode == "Translate to English",
                    )
                    st.success("Transcript ready. Review it below before generating the summary.")
                except Exception as exc:
                    st.error(f"Audio transcription failed: {exc}")

        if document_files:
            if st.button("Add File Content"):
                try:
                    for document_file in document_files:
                        extracted_text = extract_text_from_document(document_file)
                        labeled_text = f"File: {getattr(document_file, 'name', 'document')}\n{extracted_text}"
                        st.session_state.capture_transcript = append_document_to_transcript(
                            st.session_state.capture_transcript,
                            labeled_text,
                        )
                    st.success("Supporting document content added to the transcript area.")
                except Exception as exc:
                    st.error(f"Document processing failed: {exc}")

        transcript = st.text_area(
            "Transcript / Meeting Notes",
            height=260,
            placeholder="Paste your meeting transcript here or transcribe audio above...",
            key="capture_transcript",
        )

        action_col_left, action_col_right = st.columns(2)
        with action_col_left:
            run_clicked = st.button(
                "Generate Summary",
                key="capture_generate_btn",
                type="primary",
                use_container_width=True,
                disabled=not transcript.strip(),
            )
        with action_col_right:
            st.button("Clear Input", key="capture_clear_btn", on_click=clear_capture_inputs, use_container_width=True)

    resolved_activity_id = st.session_state.capture_activity_id.strip() or generate_activity_id(
        activity_category,
        meeting_date,
        st.session_state.get("meetings", []),
    )

    if run_clicked:
        progress = st.progress(0, text="Starting pipeline...")
        progress.progress(0.2, text="Reading transcript...")
        progress.progress(0.45, text="Extracting actions and decisions...")
        progress.progress(0.75, text="Preparing meeting brief...")

        try:
            attach_sources = [link_photo if link_photo and link_photo != "No Photo" else ""]
            if link_photo == "Insert link" and link_photo_url.strip():
                attach_sources.insert(0, link_photo_url.strip())
            attach_sources.extend(
                getattr(document_file, "name", "").strip() for document_file in (document_files or []) if getattr(document_file, "name", "").strip()
            )
            attach_file_value = " | ".join(part for part in attach_sources if part)
            pipeline_metadata = {
                "Category": activity_category,
                "Activity ID": resolved_activity_id,
                "Activity Title": activity_title,
                "Role": role,
                "Main Activity": main_activity,
                "Link Photo": link_photo,
                "District": district,
                "Invitation From": invitation_from,
                "Location Meeting": location_meeting,
                "Other Reps": other_reps,
                "Activity Type": activity_type,
                "Organization Type": organization_type,
                "Date From": date_from.isoformat(),
                "Date To": date_to.isoformat(),
                "Department": "" if dept_choice == dept_names[0] else dept_choice,
                "Representative Position": representative_position,
                "Representative Name": representative_name,
                "Representative Department": representative_department,
                "Link Photo URL": link_photo_url,
                "Staff Email": stfemail,
                "Supervisor Email": supemail,
                "Updated By": updated_by,
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
                "dept": dept_choice if dept_choice != dept_names[0] else "",
                "actual_cost": actual_cost,
                "meeting_date": meeting_date.isoformat(),
                "activity_category": activity_category,
                "activity_id": resolved_activity_id,
                "activity_title": activity_title,
                "role": role,
                "main_activity": main_activity,
                "link_photo": link_photo,
                "link_photo_url": link_photo_url,
                "attach_file": attach_file_value,
                "activity_type": activity_type,
                "organization_type": organization_type,
                "date_from": date_from.isoformat(),
                "date_to": date_to.isoformat(),
                "representative_position": representative_position,
                "representative_name": representative_name,
                "representative_department": representative_department,
                "district": district,
                "invitation_from": invitation_from,
                "location_meeting": location_meeting,
                "other_reps": other_reps,
                "slt_department": representative_department or ("" if dept_choice == dept_names[0] else dept_choice),
                "slt_position": representative_position,
                "slt_reps": representative_name,
                "stf_email": stfemail,
                "sup_email": supemail,
                "updated_by": updated_by,
                "activity_objective": "",
            }
        except Exception as exc:
            st.error(f"Pipeline failed: {exc}")

    if st.session_state.pending_result:
        pending = st.session_state.pending_result
        result = dict(pending["result"])
        result["title"] = result.get("title") or pending.get("activity_title", "Untitled")
        if "preview_actions" not in pending:
            pending["preview_actions"] = [
                make_action_preview_item(action, f"preview_{index}")
                for index, action in enumerate(result.get("action_items", []))
            ]
        render_summary_panel(result)

        def parse_preview_list(value: str) -> list:
            items = []
            for line in (value or "").splitlines():
                cleaned = line.strip().lstrip("-").lstrip("•").strip()
                if cleaned:
                    items.append(cleaned)
            return items

        insight_col, entity_col = st.columns([1.2, 0.8])
        with insight_col:
            st.markdown("### Decisions & Discussion")
            decisions_text = st.text_area(
                "Key Decisions",
                value="\n".join(normalize_value(item, "") for item in result.get("key_decisions", [])),
                key="preview_key_decisions",
                height=120,
                placeholder="Add one decision per line",
            )
            discussions_text = st.text_area(
                "Discussion Points",
                value="\n".join(normalize_value(item, "") for item in result.get("discussion_points", [])),
                key="preview_discussion_points",
                height=180,
                placeholder="Add one discussion point per line",
            )
            result["key_decisions"] = parse_preview_list(decisions_text)
            result["discussion_points"] = parse_preview_list(discussions_text)

        with entity_col:
            entities = result.get("nlp_pipeline", {}).get("named_entities", {})
            st.markdown("### Meeting Metadata")
            people_text = st.text_area(
                "People",
                value="\n".join(extract_entity_names(entities.get("persons", []))),
                key="preview_people",
                height=100,
                placeholder="Add one person per line",
            )
            organizations_text = st.text_area(
                "Organizations",
                value="\n".join(extract_entity_names(entities.get("organizations", []))),
                key="preview_organizations",
                height=100,
                placeholder="Add one organization per line",
            )
            dates_text = st.text_area(
                "Dates Mentioned",
                value="\n".join(extract_entity_names(entities.get("dates", []))),
                key="preview_dates",
                height=100,
                placeholder="Add one date per line",
            )
            entities["persons"] = parse_preview_list(people_text)
            entities["organizations"] = parse_preview_list(organizations_text)
            entities["dates"] = parse_preview_list(dates_text)
            entities["locations"] = []
            result.setdefault("nlp_pipeline", {})
            result["nlp_pipeline"].setdefault("named_entities", {})
            result["nlp_pipeline"]["named_entities"] = entities

        st.markdown("### Action Plan")
        preview_actions = pending.get("preview_actions", [])
        if preview_actions:
            for action in preview_actions:
                render_action_card(action, editable=True)
        else:
            st.info("No action items detected.")

        if result.get("follow_up"):
            st.markdown("**Follow-up:** Yes")

        email_preview_meeting = build_email_preview_meeting(result, pending)

        with st.expander("Prepare Email Copy"):
            render_email_copy_block(email_preview_meeting, "preview_email_copy")

        if st.button("Save Record", type="primary"):
            new_meeting = build_meeting_record(result, pending)
            st.session_state.meetings.insert(0, new_meeting)
            persist_app_data()
            st.session_state.pending_result = None
            st.success("Meeting saved to library.")
            st.rerun()


# ============================================================
# Section 8. Dashboard Tab
# ============================================================

if st.session_state.current_page == "Dashboard":
    dashboard_years = sorted(meeting_df["year"].dropna().unique().tolist(), reverse=True) if not meeting_df.empty else [date.today().year]

    def meeting_dashboard_status(meeting: dict) -> str:
        actions = meeting.get("actions", [])
        if not actions:
            return "Completed"
        statuses = [normalize_status(action) for action in actions]
        if "Overdue" in statuses:
            return "Overdue"
        if any(status in {"Pending", "In Progress"} for status in statuses):
            return "Pending"
        return "Completed"

    dashboard_meeting_records = [{"meeting": meeting, "status": meeting_dashboard_status(meeting)} for meeting in meetings]
    done_count = sum(1 for record in dashboard_meeting_records if record["status"] == "Completed")
    pending_count = sum(1 for record in dashboard_meeting_records if record["status"] == "Pending")
    completion_pct = round((done_count / len(dashboard_meeting_records)) * 100) if dashboard_meeting_records else 0

    dashboard_left, dashboard_right = st.columns([1.35, 0.85])
    with dashboard_left:
        overview_card = st.container(border=True)
        with overview_card:
            st.markdown("### Today’s Brief")
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                render_kpi_card("Meetings", str(len(meetings)), "Stored records", "#0f766e")
            with c2:
                render_kpi_card("Pending", str(pending_count), "Pending meetings", "#d97706")
            with c3:
                render_kpi_card("Done", str(done_count), "Closed or no-action meetings", "#16a34a")
            with c4:
                render_completion_ring(completion_pct)

        upcoming_card = st.container(border=True)
        with upcoming_card:
            st.markdown("### Upcoming Project")
            sort_choice = st.selectbox(
                "Sort upcoming projects",
                ["Earliest deadline", "Latest deadline"],
                key="upcoming_project_sort",
            )
            upcoming_meetings = get_upcoming_meetings(meetings, sort_order=sort_choice)
            if not upcoming_meetings:
                st.info("No upcoming projects yet.")
            else:
                for meeting in upcoming_meetings:
                    with st.expander(normalize_value(meeting.get("title"), "Untitled")):
                        st.markdown(
                            f"""
                            <div class="upcoming-item">
                                <div class="upcoming-top">
                                    <div>
                                        <div class="mini-copy">{normalize_value(meeting.get('meetingID'), 'No ID')} | {normalize_value(meeting.get('deptName') or meeting.get('department'), 'No group')}</div>
                                    </div>
                                    <div class="upcoming-date">{normalize_value(meeting.get('date'), 'No date')}</div>
                                </div>
                            </div>
                            """,
                            unsafe_allow_html=True,
                        )
                        st.markdown(f"**Summary:** {normalize_value(meeting.get('summary'), 'No summary available.')}")
                        if meeting.get("actions"):
                            st.markdown("**Action Items**")
                            for action in meeting.get("actions", []):
                                render_action_card(action)

    with dashboard_right:
        calendar_card = st.container(border=True)
        with calendar_card:
            st.markdown("### Calendar")
            calendar_year_options = sorted(
                set(dashboard_years + [date.today().year]),
                reverse=True,
            )
            month_names = list(calendar.month_name)[1:]
            current_month_name = calendar.month_name[date.today().month]
            calendar_top_left, calendar_top_right = st.columns([0.56, 0.44])
            with calendar_top_left:
                selected_calendar_month = st.selectbox(
                    "Calendar Month",
                    month_names,
                    index=month_names.index(st.session_state.get("calendar_month", current_month_name)),
                    key="calendar_month",
                    label_visibility="collapsed",
                )
            with calendar_top_right:
                default_calendar_year = st.session_state.get("calendar_year", date.today().year)
                year_index = calendar_year_options.index(default_calendar_year) if default_calendar_year in calendar_year_options else 0
                selected_calendar_year = st.selectbox(
                    "Calendar Year",
                    calendar_year_options,
                    index=year_index,
                    key="calendar_year",
                    label_visibility="collapsed",
                )
            selected_calendar_month_num = month_names.index(selected_calendar_month) + 1
            st.caption(f"{selected_calendar_month} {selected_calendar_year}")
            st.markdown(
                build_calendar_html(meetings, selected_calendar_year, selected_calendar_month_num),
                unsafe_allow_html=True,
            )
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


# ============================================================
# Section 9. Action Tracker Tab
# ============================================================

if st.session_state.current_page == "Tracker":
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
        st.caption("Showing pending meetings from the dashboard shortcut.")
    elif tracker_focus == "done":
        status_default = "Completed"
        st.caption("Showing completed meetings from the dashboard shortcut.")

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
                "keywords": " ".join(
                    [
                        normalize_value(meeting.get("title"), ""),
                        normalize_value(meeting.get("summary"), ""),
                        normalize_value(meeting.get("recaps"), ""),
                        normalize_value(meeting.get("meetingID"), ""),
                        normalize_value(meeting.get("activityId"), ""),
                        normalize_value(meeting.get("deptName"), ""),
                        normalize_value(meeting.get("department"), ""),
                        normalize_value(meeting.get("sltdepartment"), ""),
                        join_list(meeting.get("stakeholders", []), ""),
                        join_list(meeting.get("companies", []), ""),
                    ]
                ).lower(),
            }
        )

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        render_kpi_card("Saved Meetings", str(len(meeting_records)), "Stored records", "#0f766e")
    with c2:
        render_kpi_card("Pending", str(sum(1 for record in meeting_records if record["status"] == "Pending")), "Pending meetings", "#d97706")
    with c3:
        render_kpi_card("Completed", str(sum(1 for record in meeting_records if record["status"] == "Completed")), "Closed or no-action meetings", "#16a34a")
    with c4:
        completion_pct = round(
            (
                sum(1 for record in meeting_records if record["status"] == "Completed") / len(meeting_records)
            ) * 100
        ) if meeting_records else 0
        render_completion_ring(completion_pct)

    filt_left, filt_right = st.columns([1.4, 1])
    with filt_left:
        meeting_search = st.text_input(
            "Search meeting",
            placeholder="Search by event name, keyword, generated ID, or group name",
            key="tracker_meeting_search",
        )
    with filt_right:
        status_options = ["All", "Pending", "Overdue", "Completed"]
        status_index = status_options.index(status_default) if status_default in status_options else 0
        meeting_status_filter = st.selectbox("Meeting Status", status_options, index=status_index, key="tracker_meeting_status")

    if tracker_focus != "all":
        if st.button("Clear Tracker Shortcut", key="clear_tracker_shortcut"):
            st.session_state.tracker_focus = "all"
            st.session_state.tracker_meeting_status = "All"
            st.rerun()

    filtered_meetings = meeting_records
    if meeting_status_filter != "All":
        filtered_meetings = [record for record in filtered_meetings if record["status"] == meeting_status_filter]

    search_needle = meeting_search.strip().lower()
    if search_needle:
        filtered_meetings = [
            record for record in filtered_meetings
            if search_needle in record["keywords"]
        ]

    st.markdown("### Saved Meetings")
    if not filtered_meetings:
        st.info("No saved meetings match the selected search or status.")
    else:
        filtered_meetings = sorted(
            filtered_meetings,
            key=lambda record: normalize_value(record["meeting"].get("date"), "0000-00-00"),
            reverse=True,
        )
        for record in filtered_meetings:
            meeting = record["meeting"]
            status_text = record["status"]
            status_cfg = {"Pending": "#d97706", "Overdue": "#dc2626", "Completed": "#16a34a"}
            subtitle = "Pending meetings" if status_text == "Pending" else ("Past deadline meetings" if status_text == "Overdue" else "No pending action items")
            header = f"{record['title']} | {record['meeting_id'] or 'No ID'} | {normalize_value(meeting.get('date'), 'No date')}"
            with st.expander(header):
                top_left, top_right = st.columns([1.35, 0.65])
                with top_left:
                    st.markdown(
                        f"""
                        <div class="info-card">
                            <div class="mini-title">Summary</div>
                            <div class="mini-copy">{normalize_value(meeting.get('summary'), 'No summary available.')}</div>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                with top_right:
                    render_kpi_card("Status", status_text, subtitle, status_cfg[status_text])
                    render_kpi_card("Actions", str(len(meeting.get("actions", []))), "Tracked items", "#1e3a5f")

                meta_left, meta_right = st.columns(2)
                with meta_left:
                    st.markdown(
                        f"""
                        <div class="info-card">
                            <div class="mini-title">Group / Department</div>
                            <div class="mini-copy">{normalize_value(record.get('group'), 'None')}</div>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                with meta_right:
                    st.markdown(
                        f"""
                        <div class="info-card">
                            <div class="mini-title">Follow-up</div>
                            <div class="mini-copy">{'Yes' if status_text == 'Pending' else 'No'}</div>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )

                if meeting.get("keyDecisions"):
                    st.markdown("**Key Decisions**")
                    for decision in meeting["keyDecisions"]:
                        st.write(normalize_value(decision))

                if meeting.get("discussionPoints"):
                    st.markdown("**Discussion Points**")
                    for point in meeting["discussionPoints"]:
                        st.write(normalize_value(point))

                if meeting.get("actions"):
                    st.markdown("**Action Items**")
                    for action in meeting["actions"]:
                        render_action_card(action, editable=True, persist_callback=persist_app_data)
                else:
                    st.success("No action item. This meeting is considered completed.")


# ============================================================
# Section 10. Finance Tab
# ============================================================

if st.session_state.current_page == "Finance":
    st.subheader("Finance Tracker")

    total_spend = float(meeting_df["cost"].sum()) if not meeting_df.empty else 0
    total_est = float(sum(meeting.get("estimatedCost", 0) for meeting in meetings))
    over_budget = sum(1 for meeting in meetings if meeting.get("actualCost", 0) > meeting.get("estimatedCost", 0) > 0)

    c1, c2, c3 = st.columns(3)
    with c1:
        render_kpi_card("Total Spend", rm(total_spend), "Recorded across meetings", "#0f766e")
    with c2:
        render_kpi_card("Estimated Budget", rm(total_est), "AI extracted estimate", "#2563eb")
    with c3:
        render_kpi_card("Over Budget", str(over_budget), "Meetings above estimate", "#dc2626")

    if not meeting_df.empty:
        finance_years = sorted(meeting_df["year"].dropna().unique().tolist(), reverse=True)
        finance_year = st.selectbox("Finance Year", finance_years, key="finance_year_filter")
        finance_df = add_month_columns(meeting_df[meeting_df["year"] == finance_year].copy(), "date")
        finance_left, finance_right = st.columns(2)
        with finance_left:
            monthly = (
                finance_df.groupby(["month_num", "month_label"], as_index=False)["cost"]
                .sum()
                .sort_values("month_num")
            )
            fig_monthly = px.line(
                monthly,
                x="month_label",
                y="cost",
                title=f"Monthly Spend Trend ({finance_year})",
                markers=True,
                line_shape="linear",
                color_discrete_sequence=["#0f766e"],
            )
            style_plotly(fig_monthly)
            st.plotly_chart(fig_monthly, use_container_width=True)
        with finance_right:
            dept_rollup = (
                finance_df.groupby(["month_num", "month_label", "department"], as_index=False)["cost"]
                .sum()
                .sort_values("month_num")
            )
            fig_dept = px.line(
                dept_rollup,
                x="month_label",
                y="cost",
                title=f"Monthly Spend by Department ({finance_year})",
                color="department",
                markers=True,
                color_discrete_sequence=["#1e3a5f", "#0f766e", "#2563eb", "#d97706", "#7c3aed"],
            )
            style_plotly(fig_dept)
            st.plotly_chart(fig_dept, use_container_width=True)

    st.markdown("### Department Budgets")
    with st.form("add_dept"):
        dept_name = st.text_input("Department Name")
        dept_budget = st.number_input("Annual Budget (RM)", min_value=0.0, step=1000.0)
        submitted = st.form_submit_button("Add Department")
        if submitted and dept_name:
            st.session_state.departments.append({"id": uid(), "name": dept_name, "budget": dept_budget})
            persist_app_data()
            st.rerun()

    for dept in st.session_state.departments:
        dept_meetings = [meeting for meeting in meetings if meeting.get("deptId") == dept["id"]]
        spend = sum(meeting.get("actualCost", 0) for meeting in dept_meetings)
        budget = dept.get("budget", 0)
        pct = min((spend / budget) * 100 if budget else 0, 100)
        st.markdown(
            f"""
            <div class="info-card">
                <div class="mini-title">{dept['name']}</div>
                <div class="mini-copy">{rm(spend)} used of {rm(budget)}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.progress(pct / 100 if budget else 0)
