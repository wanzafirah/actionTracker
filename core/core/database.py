import json
import os
import sys
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import requests
import streamlit as st

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

try:
    import gspread
except ImportError:
    gspread = None

from meetiq_constants import (
    DATA_FILE,
    DEFAULT_DEPARTMENTS,
    DEPARTMENTS_SHEET_NAME,
    DEPARTMENT_SHEET_COLUMNS,
    HISTORY_SHEET_COLUMNS,
    HISTORY_SHEET_NAME,
    MEETINGS_SHEET_NAME,
    MEETING_SHEET_COLUMNS,
)
from meetiq_utils import extract_entity_names, first_nonempty, generate_activity_id, json_dumps_safe, json_loads_safe, load_text_list, normalize_status, normalize_value, parse_yes_no, today_str, uid, yes_no_text


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

    if page_value in {"Dashboard", "Tracker", "Capture", "History"}:
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


def parse_lines(value: str) -> list:
    items = []
    for line in (value or "").splitlines():
        cleaned = line.strip().lstrip("-").lstrip("•").strip()
        if cleaned:
            items.append(cleaned)
    return items


def normalize_chat_user_id(value: str) -> str:
    return normalize_value(value, "").strip()


def get_current_chat_user_id() -> str:
    return normalize_chat_user_id(st.session_state.get("chat_user_id", ""))


def get_chat_thread_key(user_id: str, thread_date: str, meeting_title: str, meeting_id: str) -> str:
    return " | ".join(part for part in [user_id.lower(), thread_date, meeting_title.lower(), meeting_id.lower()] if part)


def build_chat_thread_label(entry: dict) -> str:
    thread_date = normalize_value(entry.get("thread_date"), "")
    thread_title = normalize_value(entry.get("thread_title") or entry.get("meeting_title"), "General")
    meeting_id = normalize_value(entry.get("meeting_id"), "")
    label_parts = [thread_date or "No date", thread_title]
    if meeting_id:
        label_parts.append(meeting_id)
    return " | ".join(label_parts)


def parse_loaded_dataframes(meetings_df: pd.DataFrame, departments_df: pd.DataFrame, history_df: pd.DataFrame | None = None):
    meetings = []
    if not meetings_df.empty:
        for row in meetings_df.fillna("").to_dict("records"):
            meetings.append(
                {
                    "id": first_nonempty(row, "id", fallback=""),
                    "user_id": first_nonempty(row, "user_id", fallback=""),
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
            departments.append({"id": row.get("id", ""), "name": row.get("name", ""), "budget": float(row.get("budget", 0) or 0)})

    history = []
    if history_df is not None and not history_df.empty:
        for row in history_df.fillna("").to_dict("records"):
            history.append(
                {
                    "id": first_nonempty(row, "id", fallback=""),
                    "user_id": first_nonempty(row, "user_id", fallback=""),
                    "thread_key": first_nonempty(row, "thread_key", fallback=""),
                    "thread_date": first_nonempty(row, "thread_date", fallback=""),
                    "thread_title": first_nonempty(row, "thread_title", fallback=""),
                    "timestamp": first_nonempty(row, "timestamp", fallback=""),
                    "question": first_nonempty(row, "question", fallback=""),
                    "answer": first_nonempty(row, "answer", fallback=""),
                    "meeting_id": first_nonempty(row, "meeting_id", fallback=""),
                    "meeting_title": first_nonempty(row, "meeting_title", fallback=""),
                    "context": first_nonempty(row, "context", fallback=""),
                }
            )

    return meetings, departments, history


def build_meeting_rows(meetings: list) -> list:
    meetings_rows = []
    for meeting in meetings:
        attach_file_value = normalize_value(meeting.get("attachFile") or meeting.get("linkPhotoUrl") or meeting.get("linkPhoto"), "")
        meetings_rows.append({column: "" for column in MEETING_SHEET_COLUMNS})
        row = meetings_rows[-1]
        row.update(
            {
                "id": meeting.get("id", ""),
                "user_id": meeting.get("user_id", ""),
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
    return [{"id": department.get("id", ""), "name": department.get("name", ""), "budget": department.get("budget", 0)} for department in departments]


def build_history_rows(history: list) -> list:
    return [{column: entry.get(column, "") for column in HISTORY_SHEET_COLUMNS} for entry in history]


def load_excel_data():
    if not os.path.exists(DATA_FILE):
        return [], [], []
    try:
        meetings_df = pd.read_excel(DATA_FILE, sheet_name=MEETINGS_SHEET_NAME)
    except Exception:
        meetings_df = pd.DataFrame()
    try:
        departments_df = pd.read_excel(DATA_FILE, sheet_name=DEPARTMENTS_SHEET_NAME)
    except Exception:
        departments_df = pd.DataFrame()
    try:
        history_df = pd.read_excel(DATA_FILE, sheet_name=HISTORY_SHEET_NAME)
    except Exception:
        history_df = pd.DataFrame()
    return parse_loaded_dataframes(meetings_df, departments_df, history_df)


def save_excel_data(meetings: list, departments: list):
    meetings_rows = build_meeting_rows(meetings)
    departments_rows = build_department_rows(departments)
    history_rows = build_history_rows(st.session_state.get("chat_history_records", []))
    with pd.ExcelWriter(DATA_FILE, engine="openpyxl") as writer:
        pd.DataFrame(meetings_rows, columns=MEETING_SHEET_COLUMNS).to_excel(writer, sheet_name=MEETINGS_SHEET_NAME, index=False)
        pd.DataFrame(departments_rows, columns=DEPARTMENT_SHEET_COLUMNS).to_excel(writer, sheet_name=DEPARTMENTS_SHEET_NAME, index=False)
        pd.DataFrame(history_rows, columns=HISTORY_SHEET_COLUMNS).to_excel(writer, sheet_name=HISTORY_SHEET_NAME, index=False)


@st.cache_resource
def get_google_spreadsheet():
    service_account_info = get_google_service_account_info()
    target = get_google_sheet_target()
    if gspread is None or not service_account_info or not any(target.values()):
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
        history_ws = get_or_create_google_worksheet(spreadsheet, HISTORY_SHEET_NAME, HISTORY_SHEET_COLUMNS)
        return parse_loaded_dataframes(pd.DataFrame(meetings_ws.get_all_records()), pd.DataFrame(departments_ws.get_all_records()), pd.DataFrame(history_ws.get_all_records()))
    except Exception:
        return None


def save_google_sheet_data(meetings: list, departments: list):
    spreadsheet = get_google_spreadsheet()
    if spreadsheet is None:
        raise RuntimeError("Google Sheets storage is not configured.")
    meetings_ws = get_or_create_google_worksheet(spreadsheet, MEETINGS_SHEET_NAME, MEETING_SHEET_COLUMNS)
    departments_ws = get_or_create_google_worksheet(spreadsheet, DEPARTMENTS_SHEET_NAME, DEPARTMENT_SHEET_COLUMNS)
    history_ws = get_or_create_google_worksheet(spreadsheet, HISTORY_SHEET_NAME, HISTORY_SHEET_COLUMNS)
    meetings_ws.clear()
    meetings_ws.update("A1", [MEETING_SHEET_COLUMNS] + [[row.get(column, "") for column in MEETING_SHEET_COLUMNS] for row in build_meeting_rows(meetings)])
    departments_ws.clear()
    departments_ws.update("A1", [DEPARTMENT_SHEET_COLUMNS] + [[row.get(column, "") for column in DEPARTMENT_SHEET_COLUMNS] for row in build_department_rows(departments)])
    history_ws.clear()
    history_ws.update("A1", [HISTORY_SHEET_COLUMNS] + [[row.get(column, "") for column in HISTORY_SHEET_COLUMNS] for row in build_history_rows(st.session_state.get("chat_history_records", []))])


def supabase_headers(prefer: str = "") -> dict | None:
    config = get_supabase_config()
    if not config["url"] or not config["key"]:
        return None
    headers = {"apikey": config["key"], "Authorization": f"Bearer {config['key']}", "Content-Type": "application/json"}
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
    history_url = supabase_table_url(HISTORY_SHEET_NAME)
    if headers is None or meetings_url is None or departments_url is None:
        return None
    try:
        meetings_df = pd.DataFrame(requests.get(meetings_url, headers=headers, params={"select": "*"}, timeout=30).json())
        departments_df = pd.DataFrame(requests.get(departments_url, headers=headers, params={"select": "*"}, timeout=30).json())
        history_df = pd.DataFrame()
        if history_url is not None:
            try:
                history_response = requests.get(history_url, headers=headers, params={"select": "*"}, timeout=30)
                history_response.raise_for_status()
                history_df = pd.DataFrame(history_response.json())
            except Exception:
                history_df = pd.DataFrame()
        return parse_loaded_dataframes(meetings_df, departments_df, history_df)
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
    replace_supabase_table(MEETINGS_SHEET_NAME, build_meeting_rows(meetings), MEETING_SHEET_COLUMNS)
    replace_supabase_table(DEPARTMENTS_SHEET_NAME, build_department_rows(departments), DEPARTMENT_SHEET_COLUMNS)
    try:
        replace_supabase_table(HISTORY_SHEET_NAME, build_history_rows(st.session_state.get("chat_history_records", [])), HISTORY_SHEET_COLUMNS)
    except Exception:
        pass


def load_app_data():
    supabase_data = load_supabase_data()
    if supabase_data is not None:
        return supabase_data
    google_data = load_google_sheet_data()
    if google_data is not None:
        return google_data
    return load_excel_data()


def save_app_data(meetings: list, departments: list, history: list | None = None):
    if history is not None:
        st.session_state.chat_history_records = history
    if supabase_headers() is not None:
        save_supabase_data(meetings, departments)
    elif get_google_spreadsheet() is not None:
        save_google_sheet_data(meetings, departments)
    else:
        save_excel_data(meetings, departments)


def persist_app_data():
    save_app_data(st.session_state.meetings, st.session_state.departments, st.session_state.get("chat_history_records", []))


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


def make_action_preview_item(action: dict, action_id: str, status: str = "Pending") -> dict:
    department = normalize_value(action.get("department") or action.get("company"), "Not stated")
    return {
        "id": action_id,
        "text": normalize_value(action.get("text"), "Untitled action"),
        "owner": normalize_value(action.get("owner"), "Not stated"),
        "department": department,
        "company": department,
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
        actions = [{"text": normalize_value(action.get("text"), "Untitled action"), "owner": normalize_value(action.get("owner"), "Not stated"), "deadline": normalize_value(action.get("deadline"), "None"), "status": normalize_value(action.get("status"), "Pending")} for action in preview_actions]
    else:
        actions = [{"text": normalize_value(action.get("text"), "Untitled action"), "owner": normalize_value(action.get("owner"), "Not stated"), "deadline": normalize_value(action.get("deadline"), "None"), "status": "Pending"} for action in result.get("action_items", [])]
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
    selected_departments = pending.get("dept", [])
    if isinstance(selected_departments, str):
        selected_departments = [selected_departments] if selected_departments else []
    selected_department_records = [department for department_name in selected_departments if (department := find_department_by_name(department_name))]
    department_names = [department["name"] for department in selected_department_records]
    department_ids = [department["id"] for department in selected_department_records]
    department_text = ", ".join(department_names)
    manual_stakeholders = pending.get("capture_stakeholders", [])
    if isinstance(manual_stakeholders, str):
        manual_stakeholders = parse_lines(manual_stakeholders)
    extracted_stakeholders = extract_entity_names(result.get("nlp_pipeline", {}).get("named_entities", {}).get("persons", []))
    stakeholder_names = []
    for name in manual_stakeholders + extracted_stakeholders:
        cleaned = normalize_value(name, "")
        if cleaned and cleaned.lower() not in {item.lower() for item in stakeholder_names}:
            stakeholder_names.append(cleaned)
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
        "stakeholders": stakeholder_names,
        "companies": extract_entity_names(result.get("nlp_pipeline", {}).get("named_entities", {}).get("organizations", [])),
        "keyDecisions": result.get("key_decisions", []),
        "discussionPoints": result.get("discussion_points", []),
        "nlpStats": result.get("nlp_pipeline", {}),
        "transcript": pending["transcript"],
        "deptId": " | ".join(department_ids),
        "deptName": department_text,
        "department": department_text,
        "actualCost": 0,
        "budgetUsed": 0,
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
        "sltdepartment": pending.get("slt_department", department_text),
        "sltposition": pending.get("slt_position", ""),
        "sltreps": pending.get("slt_reps", ""),
        "stfemail": pending.get("stf_email", ""),
        "supemail": pending.get("sup_email", ""),
        "invitationFrom": pending.get("invitation_from", ""),
        "locationMeeting": pending.get("location_meeting", ""),
        "otherReps": pending.get("other_reps", ""),
        "updatedBy": pending.get("updated_by", ""),
        "activityObjective": pending.get("activity_objective", ""),
        "actions": ([{**action, "id": f"{meeting_id}_a{index}", "department": normalize_value(action.get("department") or action.get("company"), "Not stated")} for index, action in enumerate(preview_actions)] if preview_actions else [make_action_preview_item(action, f"{meeting_id}_a{index}") for index, action in enumerate(result.get("action_items", []))]),
    }


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
    for action in actions:
        action_lines.append(f"- {normalize_value(action.get('text'))} | Assignee: {normalize_value(action.get('owner'), 'Not stated')} | Deadline: {normalize_value(action.get('deadline'), 'None')} | Status: {normalize_status(action)}")
    decisions = meeting.get("keyDecisions", [])
    decision_lines = [f"- {normalize_value(item)}" for item in decisions] if decisions else []
    lines = [f"Dear {greeting_name},", "", "Please find below the meeting summary report for your reference.", "", summary_block, "", f"Objective: {objective}", f"Outcome: {outcome}", follow_up_line]
    if decision_lines:
        lines.extend(["", "Key Decisions:"])
        lines.extend(decision_lines)
    if action_lines:
        lines.extend(["", "Action Items:"])
        lines.extend(action_lines)
    lines.extend(["", "Regards,", closing_name])
    return "\n".join(lines)


def find_meeting_by_id(meetings: list, meeting_id: str):
    if not meeting_id:
        return None
    return next((meeting for meeting in meetings if normalize_value(meeting.get("meetingID") or meeting.get("activityId") or meeting.get("id"), "") == meeting_id), None)


def update_meeting_record(meetings: list, meeting_id: str, updates: dict) -> bool:
    target = find_meeting_by_id(meetings, meeting_id)
    if not target:
        return False
    target.update(updates)
    if "summary" in updates and "recaps" not in updates:
        target["recaps"] = updates["summary"]
    if "recaps" in updates and "summary" not in updates:
        target["summary"] = updates["recaps"]
    if "department" in updates and "deptName" not in updates:
        target["deptName"] = updates["department"]
    if "deptName" in updates and "department" not in updates:
        target["department"] = updates["deptName"]
    if "actions" in updates:
        target["actions"] = updates["actions"]
    persist_app_data()
    return True


def init_state():
    if "data_loaded" not in st.session_state:
        meetings, departments, history_records = load_app_data()
        st.session_state.meetings = meetings
        st.session_state.departments = departments
        st.session_state.chat_history_records = history_records
        st.session_state.data_loaded = True
    if "meetings" not in st.session_state:
        st.session_state.meetings = []
    if "departments" not in st.session_state:
        st.session_state.departments = []
    if "chat_history" not in st.session_state:
        st.session_state.chat_history = []
    if "chat_history_records" not in st.session_state:
        st.session_state.chat_history_records = []
    if "active_chat_user_id" not in st.session_state:
        st.session_state.active_chat_user_id = ""
    if "pending_result" not in st.session_state:
        st.session_state.pending_result = None
    if "capture_transcript" not in st.session_state:
        st.session_state.capture_transcript = ""
    if "capture_activity_id" not in st.session_state:
        st.session_state.capture_activity_id = ""
    if "current_page" not in st.session_state:
        st.session_state.current_page = "Dashboard"
    elif st.session_state.current_page not in {"Dashboard", "Tracker", "Capture", "History"}:
        st.session_state.current_page = "Dashboard"
    if "tracker_focus" not in st.session_state:
        st.session_state.tracker_focus = "all"
    if "tracker_editing_meeting_id" not in st.session_state:
        st.session_state.tracker_editing_meeting_id = ""
