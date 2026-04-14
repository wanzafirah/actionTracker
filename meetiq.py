import json
import html
import os
import re
import tempfile
import calendar
from datetime import date, datetime
from uuid import uuid4

import pandas as pd
import plotly.express as px
import requests
import streamlit as st

try:
    import whisper
except ImportError:
    whisper = None

try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None

try:
    from docx import Document
except ImportError:
    Document = None


# ============================================================
# Section 1. Configuration
# ============================================================

OLLAMA_URL = "http://localhost:11434/api/generate"
OLLAMA_MODEL = "llama3.2"
WHISPER_MODEL = "tiny"
DATA_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "meetiq_data.xlsx")


def get_ollama_url() -> str:
    secret_value = ""
    try:
        secret_value = st.secrets.get("OLLAMA_URL", "")
    except Exception:
        secret_value = ""
    return os.getenv("OLLAMA_URL", secret_value or "http://127.0.0.1:11434/api/generate")


OLLAMA_URL = get_ollama_url()

STATUSES = ["Pending", "In Progress", "Done", "Overdue", "Cancelled"]
MTG_TYPES = ["Virtual", "Physical", "Not Provided"]
CATEGORIES = ["Event", "External Meeting", "Internal Meeting", "Workshop"]
ACTIVITY_CATEGORY_OPTIONS = [
    "Internal Meeting",
    "External Meeting",
    "Workshop / Focus Group / Roundtable",
    "Courtesy Call",
    "Forum / Conference / Webinar",
    "Open Day / Career Fair / related",
    "Podcast / Interview",
    "Speaking Engagement / Sharing Session",
]
ROLE_OPTIONS = ["Organiser", "Guest / Attendees", "Speaker", "Moderator", "Panel", "Auditor"]
MAIN_ACTIVITY_OPTIONS = ["Yes", "No"]
LINK_PHOTO_OPTIONS = ["refer to COMMS repository of photos", "Insert link", "No Photo"]
ACTIVITY_TYPE_OPTIONS = ["Virtual", "Physical", "Both"]
ORGANIZATION_TYPE_OPTIONS = ["Institution", "Company"]
REPRESENTATIVE_POSITION_OPTIONS = [
    "GROUP CHIEF EXECUTIVE OFFICER",
    "GROUP CHIEF OPERATING OFFICER",
    "GROUP CHIEF STRATEGY OFFICER",
    "SENIOR VICE PRESIDENT I",
    "SENIOR VICE PRESIDENT II",
    "VICE PRESIDENT I",
    "VICE PRESIDENT II",
    "ASSISTANT VICE PRESIDENT I",
    "ASSISTANT VICE PRESIDENT II",
]
DEFAULT_DEPARTMENTS = [
    "Group client management",
    "Group communications & public relations",
    "Group information & communication technology",
    "Group finance",
    "Group government engagement & facilitation",
    "Group human resources, admin & procurement",
    "Group research & policy",
    "Group CEO liaison office",
    "Graduate & emerging talent",
    "School Talent",
    "MyMahir",
    "Women, DEI & work-life practices",
    "MyHeart facilitation",
    "MyHeart Network",
    "MYXpats operations",
    "Residence Pass talent",
]

STATUS_CFG = {
    "Pending": {"color": "#b45309", "bg": "#fde68a"},
    "In Progress": {"color": "#1d4ed8", "bg": "#bfdbfe"},
    "Done": {"color": "#166534", "bg": "#bbf7d0"},
    "Overdue": {"color": "#991b1b", "bg": "#fecaca"},
    "Cancelled": {"color": "#374151", "bg": "#d1d5db"},
}


# ============================================================
# Section 2. AI Services
# ============================================================

def call_ollama(system: str, user_msg: str, max_tokens: int = 2000) -> str:
    headers = {}
    if "ngrok" in OLLAMA_URL:
        headers["ngrok-skip-browser-warning"] = "true"

    payload = {
        "model": OLLAMA_MODEL,
        "prompt": user_msg,
        "system": system,
        "stream": False,
        "options": {
            "num_predict": max_tokens,
            "num_ctx": 2048,
            "temperature": 0.1,
            "top_p": 0.9,
        },
    }
    try:
        response = requests.post(OLLAMA_URL, json=payload, headers=headers, timeout=300)
        response.raise_for_status()
        return response.json().get("response", "")
    except requests.HTTPError as exc:
        status_code = exc.response.status_code if exc.response is not None else "unknown"
        detail = exc.response.text[:500] if exc.response is not None else str(exc)
        raise RuntimeError(f"Ollama request failed with status {status_code}. {detail}") from exc
    except requests.RequestException as exc:
        raise RuntimeError(
            f"Could not connect to Ollama at {OLLAMA_URL}. "
            "If you are deploying on Streamlit Cloud, set OLLAMA_URL to a reachable server."
        ) from exc


# ============================================================
# Section 3. Helpers
# ============================================================

def uid() -> str:
    return uuid4().hex[:9]


def today_str() -> str:
    return date.today().isoformat()


def generate_activity_id(category: str, meeting_date: date) -> str:
    category_clean = re.sub(r"[^A-Z]", "", category.upper())
    prefix = category_clean[:3] if category_clean else "ACT"
    return f"{prefix}-{meeting_date.strftime('%Y%m%d')}-{uuid4().hex[:4].upper()}"


def set_generated_activity_id():
    category = st.session_state.get("capture_activity_category", "")
    st.session_state.capture_activity_id = generate_activity_id(category, date.today())


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

    if page_value in {"Dashboard", "Capture", "Tracker", "Finance"}:
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


def days_left(deadline: str):
    try:
        return (datetime.strptime(deadline, "%Y-%m-%d").date() - date.today()).days
    except Exception:
        return None


def rm(value) -> str:
    try:
        return f"RM {float(value):,.2f}" if value else "None"
    except Exception:
        return "None"


def normalize_status(action: dict) -> str:
    current = action.get("status", "Pending")
    if current in ("Done", "Cancelled"):
        return current
    delta = days_left(action.get("deadline", ""))
    if delta is not None and delta < 0:
        return "Overdue"
    return current


def pretty_deadline(deadline: str) -> str:
    if not deadline:
        return "None"
    delta = days_left(deadline)
    if delta is None:
        return deadline
    if delta < 0:
        return f"{deadline} | {abs(delta)}d overdue"
    if delta == 0:
        return f"{deadline} | due today"
    return f"{deadline} | {delta}d left"


def pill(label: str, color: str, bg: str) -> str:
    return (
        f"<span style='display:inline-block;padding:0.3rem 0.7rem;border-radius:999px;"
        f"background:{bg};color:{color};font-weight:600;font-size:0.82rem'>{label}</span>"
    )


def join_list(items: list, fallback: str = "None") -> str:
    clean = [str(item) for item in items if str(item).strip()]
    return ", ".join(clean) if clean else fallback


def entity_text(item) -> str:
    if isinstance(item, dict):
        for key in ("text", "name", "title", "value"):
            value = item.get(key)
            if value:
                return str(value).strip()
        return ""
    return str(item).strip()


def entity_confidence(item) -> str:
    if isinstance(item, dict) and item.get("confidence") is not None:
        try:
            return f" ({float(item['confidence']):.1f})"
        except Exception:
            return ""
    return ""


def render_entity_list(items: list, fallback: str = "None") -> str:
    cleaned = []
    for item in items or []:
        text = entity_text(item)
        if text:
            cleaned.append(f"{text}{entity_confidence(item)}")
    return ", ".join(cleaned) if cleaned else fallback


def extract_entity_names(items: list) -> list:
    names = []
    for item in items or []:
        text = entity_text(item)
        if text:
            names.append(text)
    return names


def action_belongs_to_talentcorp(action: dict) -> bool:
    company = normalize_value(action.get("company"), "").strip().lower()
    owner = normalize_value(action.get("owner"), "").strip().lower()
    text = normalize_value(action.get("text"), "").strip().lower()
    allowed = ("", "none", "not stated", "internal", "talentcorp", "talent corp")
    if company in allowed:
        return True
    if "talentcorp" in company or "talent corp" in company:
        return True
    if any(token in company for token in ("external", "partner", "company", "university", "college", "school")):
        return False
    if "talentcorp" in owner or "talent corp" in owner:
        return True
    if any(token in text for token in ("talentcorp", "our team", "internal team", "ce team", "comms", "slt")):
        return True
    return False


def filter_talentcorp_actions(actions: list) -> list:
    return [action for action in actions or [] if action_belongs_to_talentcorp(action)]


def normalize_value(value, fallback: str = "None") -> str:
    if value is None:
        return fallback
    if isinstance(value, str):
        return value.strip() or fallback
    if isinstance(value, dict):
        for key in ("text", "title", "name", "decision", "point", "summary", "value"):
            if key in value and str(value[key]).strip():
                return str(value[key]).strip()
        parts = [f"{k}: {v}" for k, v in value.items() if str(v).strip()]
        return "; ".join(parts) if parts else fallback
    if isinstance(value, list):
        parts = [normalize_value(item, "") for item in value]
        parts = [part for part in parts if part]
        return ", ".join(parts) if parts else fallback
    return str(value).strip() or fallback


def html_lines(items, fallback: str = "None") -> str:
    if not items:
        return fallback
    rendered = [normalize_value(item, "") for item in items]
    rendered = [item for item in rendered if item]
    return "<br>".join(rendered) if rendered else fallback


def style_plotly(fig, height: int = 360):
    fig.update_layout(
        template="plotly_white",
        height=height,
        paper_bgcolor="#ffffff",
        plot_bgcolor="#ffffff",
        font=dict(color="#0f172a", size=13),
        title_font=dict(color="#0f172a", size=18),
        margin=dict(l=16, r=16, t=56, b=16),
    )
    fig.update_xaxes(showgrid=False, color="#334155")
    fig.update_yaxes(gridcolor="#e2e8f0", zerolinecolor="#e2e8f0", color="#334155")
    return fig


@st.cache_resource
def get_local_whisper_model():
    if whisper is None:
        raise RuntimeError("Local Whisper is not installed. Install it with `pip install openai-whisper`.")
    return whisper.load_model(WHISPER_MODEL)


def transcribe_audio_file(uploaded_file, translate_to_english: bool = True) -> str:
    model = get_local_whisper_model()
    file_name = getattr(uploaded_file, "name", "meeting_audio.wav")
    file_bytes = uploaded_file.getvalue()
    if not file_bytes:
        raise RuntimeError("Audio file is empty.")

    suffix = os.path.splitext(file_name)[1] or ".wav"
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as temp_file:
        temp_file.write(file_bytes)
        temp_path = temp_file.name

    try:
        result = model.transcribe(
            temp_path,
            task="translate" if translate_to_english else "transcribe",
            fp16=False,
        )
        return (result.get("text") or "").strip()
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)


def dataframe_to_meeting_text(frame: pd.DataFrame, row_limit: int = 40) -> str:
    frame = frame.fillna("")
    if frame.empty:
        return ""

    records = []
    for row_index, row in frame.head(row_limit).iterrows():
        fields = []
        for column in frame.columns:
            value = str(row[column]).strip()
            if value:
                fields.append(f"{column}: {value}")
        if fields:
            records.append(f"Row {row_index + 1}: " + " | ".join(fields))
    return "\n".join(records)


def extract_text_from_document(uploaded_file) -> str:
    file_name = getattr(uploaded_file, "name", "document").lower()

    if file_name.endswith(".pdf"):
        if PdfReader is None:
            raise RuntimeError("PDF support requires `pypdf`. Install it with `pip install pypdf`.")
        reader = PdfReader(uploaded_file)
        pages = [page.extract_text() or "" for page in reader.pages]
        return "\n".join(page.strip() for page in pages if page.strip())

    if file_name.endswith(".docx"):
        if Document is None:
            raise RuntimeError("Word support requires `python-docx`. Install it with `pip install python-docx`.")
        doc = Document(uploaded_file)
        return "\n".join(paragraph.text.strip() for paragraph in doc.paragraphs if paragraph.text.strip())

    if file_name.endswith((".xlsx", ".xls")):
        sheets = pd.read_excel(uploaded_file, sheet_name=None)
        chunks = []
        for sheet_name, frame in sheets.items():
            chunks.append(f"Sheet: {sheet_name}")
            chunks.append(dataframe_to_meeting_text(frame))
        return "\n\n".join(chunk for chunk in chunks if chunk.strip())

    if file_name.endswith(".csv"):
        frame = pd.read_csv(uploaded_file)
        return dataframe_to_meeting_text(frame)

    raise RuntimeError("Unsupported document format. Use PDF, DOCX, XLSX, XLS, or CSV.")


def json_dumps_safe(value) -> str:
    return json.dumps(value if value is not None else {}, ensure_ascii=False)


def json_loads_safe(value, fallback):
    if value in ("", None) or (isinstance(value, float) and pd.isna(value)):
        return fallback
    try:
        return json.loads(value)
    except Exception:
        return fallback


def first_nonempty(row: dict, *keys: str, fallback=""):
    for key in keys:
        value = row.get(key, "")
        if value in ("", None):
            continue
        if isinstance(value, float) and pd.isna(value):
            continue
        return value
    return fallback


def parse_yes_no(value) -> bool:
    if isinstance(value, str):
        normalized = value.strip().lower()
        if normalized in ("yes", "y", "true", "1"):
            return True
        if normalized in ("no", "n", "false", "0"):
            return False
    return bool(value)


def yes_no_text(value) -> str:
    return "Yes" if parse_yes_no(value) else "No"


def load_text_list(value) -> list:
    if value in ("", None) or (isinstance(value, float) and pd.isna(value)):
        return []
    if isinstance(value, list):
        return [str(item).strip() for item in value if str(item).strip()]
    if isinstance(value, str):
        parsed = json_loads_safe(value, None)
        if isinstance(parsed, list):
            return [str(item).strip() for item in parsed if str(item).strip()]
        return [part.strip() for part in re.split(r"[;,]\s*", value) if part.strip()]
    text = str(value).strip()
    return [text] if text else []


def load_excel_data():
    if not os.path.exists(DATA_FILE):
        return [], []

    try:
        meetings_df = pd.read_excel(DATA_FILE, sheet_name="meetings")
    except Exception:
        meetings_df = pd.DataFrame()

    try:
        departments_df = pd.read_excel(DATA_FILE, sheet_name="departments")
    except Exception:
        departments_df = pd.DataFrame()

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


def save_excel_data(meetings: list, departments: list):
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

    departments_rows = [
        {
            "id": department.get("id", ""),
            "name": department.get("name", ""),
            "budget": department.get("budget", 0),
        }
        for department in departments
    ]

    with pd.ExcelWriter(DATA_FILE, engine="openpyxl") as writer:
        pd.DataFrame(meetings_rows).to_excel(writer, sheet_name="meetings", index=False)
        pd.DataFrame(departments_rows).to_excel(writer, sheet_name="departments", index=False)


def persist_app_data():
    save_excel_data(st.session_state.meetings, st.session_state.departments)


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
        "actions": [
            {
                "text": normalize_value(action.get("text"), "Untitled action"),
                "owner": normalize_value(action.get("owner"), "Not stated"),
                "deadline": normalize_value(action.get("deadline"), "None"),
                "status": "Pending",
            }
            for action in result.get("action_items", [])
        ],
    }


def build_meeting_record(result: dict, pending: dict) -> dict:
    meeting_id = uid()
    department = find_department_by_name(pending["dept"])
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
        "actions": [
            make_action_preview_item(action, f"{meeting_id}_a{index}")
            for index, action in enumerate(result.get("action_items", []))
        ],
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

def render_kpi_card(title: str, value: str, subtitle: str, accent: str) -> None:
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">{title}</div>
            <div class="kpi-value" style="color:{accent}">{value}</div>
            <div class="kpi-subtitle">{subtitle}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_kpi_link_card(title: str, value: str, subtitle: str, accent: str, page: str, focus: str) -> None:
    href = f"?page={page}&focus={focus}"
    st.markdown(
        f"""
        <a class="card-link" href="{href}">
            <div class="kpi-card clickable-card">
                <div class="kpi-label">{title}</div>
                <div class="kpi-value" style="color:{accent}">{value}</div>
                <div class="kpi-subtitle">{subtitle}</div>
            </div>
        </a>
        """,
        unsafe_allow_html=True,
    )


def render_completion_ring(percent: int) -> None:
    safe_percent = max(0, min(int(percent), 100))
    st.markdown(
        f"""
        <div class="completion-card">
            <div class="kpi-label">Completion</div>
            <div class="completion-wrap">
                <div class="completion-ring" style="--pct:{safe_percent};">
                    <div class="completion-inner">{safe_percent}%</div>
                </div>
            </div>
            <div class="kpi-subtitle">Action completeness</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_completion_link_ring(percent: int, page: str, focus: str) -> None:
    safe_percent = max(0, min(int(percent), 100))
    href = f"?page={page}&focus={focus}"
    st.markdown(
        f"""
        <a class="card-link" href="{href}">
            <div class="completion-card clickable-card">
                <div class="kpi-label">Completion</div>
                <div class="completion-wrap">
                    <div class="completion-ring" style="--pct:{safe_percent};">
                        <div class="completion-inner">{safe_percent}%</div>
                    </div>
                </div>
                <div class="kpi-subtitle">Action completeness</div>
            </div>
        </a>
        """,
        unsafe_allow_html=True,
    )


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


def render_action_card(action: dict, editable: bool = False) -> None:
    status = normalize_status(action)
    cfg = STATUS_CFG.get(status, STATUS_CFG["Pending"])
    owner = normalize_value(action.get("owner"), "Not stated")
    company = normalize_value(action.get("company"), "Internal")
    suggestion = action.get("suggestion", "")
    st.markdown(
        f"""
        <div class="action-card">
            <div class="action-top">
                <div class="action-title">{normalize_value(action.get('text'), 'Untitled action')}</div>
                {pill(status, cfg['color'], cfg['bg'])}
            </div>
            <div class="action-meta">Assignee: {owner} | Company: {company} | Deadline: {pretty_deadline(normalize_value(action.get('deadline'), 'None'))}</div>
            <div class="action-subtle">{normalize_value(suggestion, 'No next-step suggestion generated.')}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    if editable:
        current = action.get("status", "Pending")
        new_status = st.selectbox(
            "Update status",
            STATUSES,
            index=STATUSES.index(current) if current in STATUSES else 0,
            key=f"status_{action['id']}",
            label_visibility="collapsed",
        )
        if new_status != current:
            action["status"] = new_status
            persist_app_data()


def render_chat_bubble(role: str, text: str) -> None:
    safe_text = html.escape(str(text)).replace("\n", "<br>")
    st.markdown(
        f"""
        <div class="chat-bubble {role}">
            {safe_text}
        </div>
        """,
        unsafe_allow_html=True,
    )


def build_meeting_dataframe(meetings: list) -> pd.DataFrame:
    if not meetings:
        return pd.DataFrame()
    rows = []
    for meeting in meetings:
        nlp_stats = meeting.get("nlpStats", {})
        rows.append(
            {
                "id": meeting["id"],
                "title": meeting["title"],
                "date": meeting["date"],
                "month": pd.to_datetime(meeting["date"]).strftime("%Y-%m"),
                "year": pd.to_datetime(meeting["date"]).year,
                "type": meeting.get("type", "Not Provided"),
                "category": meeting.get("category", "Internal Meeting"),
                "follow_up": bool(meeting.get("followUp")),
                "actions_count": len(meeting.get("actions", [])),
                "decisions_count": len(meeting.get("keyDecisions", [])),
                "discussion_count": len(meeting.get("discussionPoints", [])),
                "cost": meeting.get("actualCost", 0),
                "token_count": nlp_stats.get("token_count", 0),
                "sentence_count": nlp_stats.get("sentence_count", 0),
                "department": meeting.get("deptName", "") or "Unassigned",
            }
        )
    return pd.DataFrame(rows)


def build_action_dataframe(meetings: list) -> pd.DataFrame:
    rows = []
    for meeting in meetings:
        for action in meeting.get("actions", []):
            rows.append(
                {
                    "id": action["id"],
                    "meeting_id": meeting["id"],
                    "meeting_title": meeting["title"],
                    "meeting_date": meeting["date"],
                    "text": normalize_value(action.get("text"), "Untitled action"),
                    "owner": normalize_value(action.get("owner"), "Not stated"),
                    "company": normalize_value(action.get("company"), "Internal"),
                    "deadline": normalize_value(action.get("deadline"), "None"),
                    "status": normalize_status(action),
                    "priority": action.get("priority", "Medium"),
                }
            )
    return pd.DataFrame(rows)


def add_month_columns(df: pd.DataFrame, date_column: str) -> pd.DataFrame:
    if df.empty:
        return df.copy()
    output = df.copy()
    parsed_dates = pd.to_datetime(output[date_column], errors="coerce")
    output["year"] = parsed_dates.dt.year
    output["month_num"] = parsed_dates.dt.month
    output["month_label"] = parsed_dates.dt.strftime("%b")
    return output


def search_meetings(meetings: list, query: str) -> list:
    needle = query.strip().lower()
    if not needle:
        return meetings[:6]

    matches = []
    for meeting in meetings:
        haystack_parts = [
            meeting.get("title", ""),
            meeting.get("summary", ""),
            meeting.get("meetingID", ""),
            meeting.get("activityId", ""),
            meeting.get("deptName", ""),
            meeting.get("department", ""),
            meeting.get("sltdepartment", ""),
            join_list(meeting.get("stakeholders", []), ""),
            join_list(meeting.get("companies", []), ""),
        ]
        haystack = " ".join(haystack_parts).lower()
        if needle in haystack:
            matches.append(meeting)
    return matches


def get_upcoming_meetings(meetings: list, limit: int = 4) -> list:
    dated_meetings = []
    for meeting in meetings:
        try:
            meeting_date = datetime.strptime(str(meeting.get("date", "")), "%Y-%m-%d").date()
            dated_meetings.append((meeting_date, meeting))
        except Exception:
            continue

    upcoming = [item for item in dated_meetings if item[0] >= date.today()]
    selected = sorted(upcoming, key=lambda item: item[0])[:limit]
    if not selected:
        selected = sorted(dated_meetings, key=lambda item: item[0], reverse=True)[:limit]
    return [meeting for _, meeting in selected]


def build_calendar_html(meetings: list, year: int, month: int) -> str:
    cal = calendar.Calendar(firstweekday=0)
    meeting_days = set()
    for meeting in meetings:
        try:
            parsed = datetime.strptime(str(meeting.get("date", "")), "%Y-%m-%d").date()
        except Exception:
            continue
        if parsed.year == year and parsed.month == month:
            meeting_days.add(parsed.day)

    weeks = cal.monthdayscalendar(year, month)
    day_labels = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
    header = "".join(f"<div class='calendar-day-label'>{label}</div>" for label in day_labels)
    cells = []
    today = date.today()
    for week in weeks:
        for day_num in week:
            classes = ["calendar-day"]
            label = "" if day_num == 0 else str(day_num)
            if day_num == today.day and month == today.month and year == today.year:
                classes.append("today")
            if day_num in meeting_days:
                classes.append("has-event")
            if day_num == 0:
                classes.append("empty")
            cells.append(f"<div class='{' '.join(classes)}'>{label}</div>")

    return f"""
    <div class="calendar-widget">
        <div class="calendar-grid calendar-head">{header}</div>
        <div class="calendar-grid">{''.join(cells)}</div>
    </div>
    """


def render_search_result_card(meeting: dict) -> None:
    st.markdown(
        f"""
        <div class="search-result-card">
            <div class="search-result-top">
                <div>
                    <div class="mini-title">{normalize_value(meeting.get('title'), 'Untitled meeting')}</div>
                    <div class="mini-copy">{normalize_value(meeting.get('meetingID'), 'No ID')} | {normalize_value(meeting.get('deptName') or meeting.get('department'), 'No group')} | {normalize_value(meeting.get('date'), 'No date')}</div>
                </div>
                <div class="result-pill">{'Yes' if meeting.get('followUp') else 'No'} follow-up</div>
            </div>
            <div class="mini-copy">{normalize_value(meeting.get('summary'), 'No summary available.')}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    actions = meeting.get("actions", [])
    if actions:
        st.markdown("**Tasks**")
        for action in actions:
            render_action_card(action)


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

def render_summary_panel(result: dict) -> None:
    nlp = result.get("nlp_pipeline", {})
    people_count = len(nlp.get("named_entities", {}).get("persons", []))
    action_count = len(result.get("action_items", []))
    decision_count = len(result.get("key_decisions", []))

    st.markdown(
        f"""
        <div class="hero-panel">
            <div class="hero-badge">Executive Meeting Brief</div>
            <h2>{normalize_value(result.get("title"), "Untitled meeting")}</h2>
            <p>{normalize_value(result.get("summary"), "No summary generated.")}</p>
            <div class="hero-grid">
                <div><strong>Objective</strong><br>{normalize_value(result.get("objective"), "Not provided")}</div>
                <div><strong>Follow-up</strong><br>{'Yes' if result.get('follow_up') else 'No'}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    k1, k2, k3 = st.columns(3)
    with k1:
        render_kpi_card("Action Items", str(action_count), "Extracted tasks", "#0f766e")
    with k2:
        render_kpi_card("Decisions", str(decision_count), "Decision signals", "#2563eb")
    with k3:
        render_kpi_card("People", str(people_count), "People involved", "#d97706")


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
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        match = re.search(r"\{.*\}", cleaned, re.DOTALL)
        if match:
            return json.loads(match.group(0))
        raise


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
    result = extract_json(raw)
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
    result["follow_up_reason"] = ""
    filtered_actions = filter_talentcorp_actions(result.get("action_items", []))
    result["action_items"] = filtered_actions
    result.setdefault("classification", {})
    result["classification"]["action_items_count"] = len(filtered_actions)
    if not filtered_actions:
        result["follow_up"] = False
    return result


def chat_with_meetings(question: str, meetings: list) -> str:
    meeting_blocks = []
    for meeting in meetings[:10]:
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
                    f"Summary: {meeting.get('summary', '')}",
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
    return call_ollama(system, user_msg, max_tokens=400)


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
    .card-link {
        text-decoration: none !important;
        display: block;
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
    .calendar-day.has-event {
        box-shadow: inset 0 0 0 2px #4f46e5;
        color: #312e81;
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
        meetings, departments = load_excel_data()
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
sync_page_from_query()
seed_default_departments()

meetings = st.session_state.meetings
meeting_df = build_meeting_dataframe(meetings)
action_df = build_action_dataframe(meetings)

with st.sidebar:
    st.markdown(
        """
        <div class="sidebar-title">AI-Powered Meeting Insight Generator &amp; Action Tracker</div>
        <div class="sidebar-subtitle">for Talentcorp by z</div>
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
        st.markdown("### Activity Details")
        act_left, act_right = st.columns(2)
        with act_left:
            activity_category = st.selectbox("Category", ACTIVITY_CATEGORY_OPTIONS, key="capture_activity_category")
            activity_id = st.text_input("Activity ID", key="capture_activity_id", placeholder="Generate or enter activity ID")
            st.button("Generate Activity ID", key="generate_activity_id_btn", on_click=set_generated_activity_id)
            st.button("Clear Activity ID", key="clear_activity_id_btn", on_click=clear_generated_activity_id)
            district = st.text_input("District", key="capture_district", placeholder="Enter district if available")
            invitation_from = st.text_input("Invitation From", key="capture_invitation_from", placeholder="Who invited this meeting?")
            location_meeting = st.text_input("Location Meeting", key="capture_location_meeting", placeholder="Meeting location")
            other_reps = st.text_input("Other Reps", key="capture_other_reps", placeholder="Other representatives or attendees")
        with act_right:
            activity_title = st.text_input("Title", key="capture_activity_title", placeholder="Enter meeting or activity title")
            role = st.selectbox("Role", ROLE_OPTIONS, key="capture_role")
            main_activity = st.selectbox("Main Activity", MAIN_ACTIVITY_OPTIONS, key="capture_main_activity")
            link_photo = st.selectbox("Attach File", LINK_PHOTO_OPTIONS, key="capture_link_photo")
            link_photo_url = ""
            if link_photo == "Insert link":
                link_photo_url = st.text_input(
                    "Attachment Link",
                    key="capture_link_photo_url",
                    placeholder="Paste the photo URL here",
                )
            activity_type = st.selectbox("Activity Type", ACTIVITY_TYPE_OPTIONS, key="capture_activity_type")
            organization_type = st.selectbox("Organization Type", ORGANIZATION_TYPE_OPTIONS, key="capture_organization_type")
            date_from = st.date_input("Date From", value=date.today(), key="capture_date_from")
            date_to = st.date_input("Date To", value=date.today(), key="capture_date_to")
            actual_cost = st.number_input("Actual Cost (RM)", min_value=0.0, step=50.0)
            dept_names = get_department_options()
            dept_choice = st.selectbox("Department", dept_names)
            meeting_date = st.date_input("Meeting Date", value=date.today())
            st.markdown("### SLT Representative Details")
            representative_position = st.selectbox("Representative Position", REPRESENTATIVE_POSITION_OPTIONS, key="capture_representative_position")
            representative_name = st.text_input("Representative Name", key="capture_representative_name", placeholder="Enter representative name")
            representative_department = st.text_input(
                "Representative Department",
                value=dept_choice if dept_choice != dept_names[0] else "",
                key="capture_representative_department",
            )
            stfemail = st.text_input("Staff Email", key="capture_stfemail", placeholder="Staff email")
            supemail = st.text_input("Supervisor Email", key="capture_supemail", placeholder="Supervisor email")
            updated_by = st.text_input("Updated By", key="capture_updated_by", placeholder="Last updated by")

    transcript_box = st.container(border=True)
    with transcript_box:
        st.markdown("### Audio Intake")
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
                type="primary",
                use_container_width=True,
                disabled=not transcript.strip(),
            )
        with action_col_right:
            st.button("Clear Input", on_click=clear_capture_inputs, use_container_width=True)

    resolved_activity_id = st.session_state.capture_activity_id.strip() or generate_activity_id(
        activity_category or activity_type or "ACT",
        date.today(),
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
        render_summary_panel(result)

        insight_col, entity_col = st.columns([1.2, 0.8])
        with insight_col:
            st.markdown("### Decisions & Discussion")
            decisions = result.get("key_decisions", [])
            discussions = result.get("discussion_points", [])
            st.markdown(
                f"""
                <div class="info-card">
                    <div class="mini-title">Key Decisions</div>
                    <div class="mini-copy">{html_lines(decisions, 'No decisions extracted.')}</div>
                </div>
                <div class="info-card">
                    <div class="mini-title">Discussion Points</div>
                    <div class="mini-copy">{html_lines(discussions, 'No discussion points extracted.')}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        with entity_col:
            entities = result.get("nlp_pipeline", {}).get("named_entities", {})
            st.markdown("### Meeting Metadata")
            st.markdown(
                f"""
                <div class="info-card">
                    <div class="mini-title">People</div>
                    <div class="mini-copy">{render_entity_list(entities.get('persons', []))}</div>
                </div>
                <div class="info-card">
                    <div class="mini-title">Organizations</div>
                    <div class="mini-copy">{render_entity_list(entities.get('organizations', []))}</div>
                </div>
                <div class="info-card">
                    <div class="mini-title">Dates & Locations</div>
                    <div class="mini-copy">{render_entity_list(entities.get('dates', []))} | {render_entity_list(entities.get('locations', []))}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        st.markdown("### Action Plan")
        actions = result.get("action_items", [])
        if actions:
            for index, item in enumerate(actions):
                render_action_card(make_action_preview_item(item, f"preview_{index}"))
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
    st.subheader("Dashboard")

    dashboard_years = sorted(meeting_df["year"].dropna().unique().tolist(), reverse=True) if not meeting_df.empty else [date.today().year]
    selected_year = st.selectbox("Dashboard Year", dashboard_years, key="dashboard_year_filter")

    done_count = int((action_df["status"] == "Done").sum()) if not action_df.empty else 0
    overdue_count = int((action_df["status"] == "Overdue").sum()) if not action_df.empty else 0
    open_count = int(len(action_df[action_df["status"].isin(["Pending", "In Progress", "Overdue"])])) if not action_df.empty else 0
    completion_pct = round((done_count / len(action_df)) * 100) if not action_df.empty and len(action_df) else 0

    dashboard_search = st.text_input(
        "Search meetings",
        placeholder="Search event name, ID, or group name...",
        label_visibility="collapsed",
        key="dashboard_search",
    )

    dashboard_left, dashboard_right = st.columns([1.35, 0.85])
    with dashboard_left:
        overview_card = st.container(border=True)
        with overview_card:
            st.markdown("### Today at a Glance")
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                render_kpi_card("Meetings", str(len(meetings)), "Stored records", "#0f766e")
            with c2:
                render_kpi_link_card("Open Tasks", str(open_count), "Pending follow-up", "#d97706", "Tracker", "open")
            with c3:
                render_kpi_link_card("Done", str(done_count), "Completed actions", "#16a34a", "Tracker", "done")
            with c4:
                render_completion_link_ring(completion_pct, "Tracker", "done")

        if not meeting_df.empty:
            year_df = add_month_columns(meeting_df[meeting_df["year"] == selected_year].copy(), "date")
            year_actions_df = add_month_columns(
                action_df[action_df["meeting_date"].astype(str).str.startswith(str(selected_year))].copy() if not action_df.empty else pd.DataFrame(),
                "meeting_date",
            )
            insights_card = st.container(border=True)
            with insights_card:
                st.markdown("### Dashboard Insights")
                overview_col, type_col = st.columns(2)
                with overview_col:
                    meetings_monthly = (
                        year_df.groupby(["month_num", "month_label"], as_index=False).size().rename(columns={"size": "count"})
                        if not year_df.empty else pd.DataFrame(columns=["month_num", "month_label", "count"])
                    )
                    done_monthly = (
                        year_actions_df[year_actions_df["status"] == "Done"].groupby(["month_num", "month_label"], as_index=False).size().rename(columns={"size": "count"})
                        if not year_actions_df.empty else pd.DataFrame(columns=["month_num", "month_label", "count"])
                    )
                    overdue_monthly = (
                        year_actions_df[year_actions_df["status"] == "Overdue"].groupby(["month_num", "month_label"], as_index=False).size().rename(columns={"size": "count"})
                        if not year_actions_df.empty else pd.DataFrame(columns=["month_num", "month_label", "count"])
                    )
                    overview_df = pd.concat(
                        [
                            meetings_monthly.assign(metric="Meetings"),
                            done_monthly.assign(metric="Done"),
                            overdue_monthly.assign(metric="Overdue"),
                        ],
                        ignore_index=True,
                    ).sort_values("month_num")
                    fig_overview = px.bar(
                        overview_df,
                        x="month_label",
                        y="count",
                        color="metric",
                        title=f"Monthly Activity ({selected_year})",
                        barmode="group",
                        color_discrete_map={"Meetings": "#1e3a5f", "Done": "#16a34a", "Overdue": "#dc2626"},
                    )
                    style_plotly(fig_overview, height=320)
                    st.plotly_chart(fig_overview, use_container_width=True)
                with type_col:
                    type_rollup = (
                        year_df.groupby(["month_num", "month_label", "type"], as_index=False).size().rename(columns={"size": "count"}).sort_values("month_num")
                        if not year_df.empty else pd.DataFrame(columns=["month_num", "month_label", "type", "count"])
                    )
                    fig_type = px.bar(
                        type_rollup,
                        x="month_label",
                        y="count",
                        color="type",
                        title=f"Meeting Type ({selected_year})",
                        barmode="group",
                        color_discrete_sequence=["#4f46e5", "#0f766e", "#94a3b8"],
                    )
                    style_plotly(fig_type, height=320)
                    st.plotly_chart(fig_type, use_container_width=True)

                spend_rollup = (
                    year_df.groupby(["month_num", "month_label", "department"], as_index=False)["cost"].sum().sort_values("month_num")
                    if not year_df.empty else pd.DataFrame(columns=["month_num", "month_label", "department", "cost"])
                )
                fig_spend = px.bar(
                    spend_rollup,
                    x="month_label",
                    y="cost",
                    color="department",
                    barmode="group",
                    title=f"Monthly Budget Spend by Department ({selected_year})",
                    color_discrete_sequence=["#1e3a5f", "#0f766e", "#2563eb", "#d97706", "#7c3aed"],
                )
                style_plotly(fig_spend, height=340)
                st.plotly_chart(fig_spend, use_container_width=True)

        results_card = st.container(border=True)
        with results_card:
            st.markdown("### Search Results")
            if not meetings:
                st.info("No meetings stored yet. Capture one to start building the dashboard.")
            else:
                matched_meetings = search_meetings(meetings, dashboard_search)
                if not matched_meetings:
                    st.info("No matching meetings found.")
                else:
                    for meeting in matched_meetings[:5]:
                        render_search_result_card(meeting)

    with dashboard_right:
        calendar_card = st.container(border=True)
        with calendar_card:
            st.markdown("### Calendar")
            st.caption(f"{calendar.month_name[date.today().month]} {date.today().year}")
            st.markdown(build_calendar_html(meetings, date.today().year, date.today().month), unsafe_allow_html=True)

        upcoming_card = st.container(border=True)
        with upcoming_card:
            st.markdown("### Upcoming Project")
            upcoming_meetings = get_upcoming_meetings(meetings)
            if not upcoming_meetings:
                st.info("No upcoming projects yet.")
            else:
                for meeting in upcoming_meetings:
                    st.markdown(
                        f"""
                        <div class="upcoming-item">
                            <div class="upcoming-top">
                                <div>
                                    <div class="mini-title">{normalize_value(meeting.get('title'), 'Untitled')}</div>
                                    <div class="mini-copy">{normalize_value(meeting.get('meetingID'), 'No ID')} | {normalize_value(meeting.get('deptName') or meeting.get('department'), 'No group')}</div>
                                </div>
                                <div class="upcoming-date">{normalize_value(meeting.get('date'), 'No date')}</div>
                            </div>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )

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

    if action_df.empty:
        st.info("No action items yet.")
    else:
        tracker_focus = st.session_state.get("tracker_focus", "all")
        if tracker_focus == "open":
            st.caption("Showing open items from the dashboard shortcut.")
        elif tracker_focus == "done":
            st.caption("Showing completed items from the dashboard shortcut.")

        c1, c2, c3, c4 = st.columns(4)
        with c1:
            render_kpi_card("Open Actions", str(len(action_df[action_df["status"].isin(["Pending", "In Progress", "Overdue"])])), "Work still active", "#0f766e")
        with c2:
            render_kpi_card("Overdue", str(len(action_df[action_df["status"] == "Overdue"])), "Needs attention", "#dc2626")
        with c3:
            render_kpi_card("Owners", str(action_df["owner"].nunique()), "Distinct assignees", "#2563eb")
        with c4:
            render_kpi_card("Companies", str(action_df["company"].nunique()), "Cross-org tasks", "#7c3aed")

        filt_left, filt_mid, filt_right = st.columns([1, 1, 1.2])
        with filt_left:
            company_filter = st.selectbox("Company", ["All"] + sorted(action_df["company"].dropna().unique().tolist()))
        with filt_mid:
            status_filter = st.selectbox("Status", ["All"] + STATUSES)
        with filt_right:
            action_search = st.text_input("Search action", placeholder="Search by task, assignee, or meeting")

        filtered_actions = filter_action_records(action_df, company_filter, status_filter, action_search)
        if tracker_focus == "open":
            filtered_actions = filtered_actions[filtered_actions["status"].isin(["Pending", "In Progress", "Overdue"])]
        elif tracker_focus == "done":
            filtered_actions = filtered_actions[filtered_actions["status"] == "Done"]

        if tracker_focus != "all":
            if st.button("Clear Tracker Shortcut", key="clear_tracker_shortcut"):
                st.session_state.tracker_focus = "all"
                st.rerun()

        st.markdown("### Action Queue")
        if filtered_actions.empty:
            st.info("No action items match the selected filters.")
        else:
            filtered_actions = filtered_actions.sort_values(["meeting_date", "meeting_title"], ascending=[False, True])
            for meeting_date, day_group in filtered_actions.groupby("meeting_date", sort=False):
                st.markdown(f"<div class='date-group'>{meeting_date}</div>", unsafe_allow_html=True)
                for record in day_group.to_dict("records"):
                    raw_action = next(
                        (
                            action
                            for meeting in meetings
                            if meeting["id"] == record["meeting_id"]
                            for action in meeting.get("actions", [])
                            if action["id"] == record["id"]
                        ),
                        None,
                    )
                    if raw_action:
                        st.caption(record["meeting_title"])
                        render_action_card(raw_action)


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
            fig_dept = px.bar(
                dept_rollup,
                x="month_label",
                y="cost",
                title=f"Monthly Spend by Department ({finance_year})",
                color="department",
                barmode="group",
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
