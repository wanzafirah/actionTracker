import os


OLLAMA_MODEL = "llama3.2:latest"
WHISPER_MODEL = "tiny"
DATA_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "meetiq_data.xlsx")
MEETINGS_SHEET_NAME = "meetings"
DEPARTMENTS_SHEET_NAME = "departments"
HISTORY_SHEET_NAME = "history"

MEETING_SHEET_COLUMNS = [
    "id",
    "user_id",
    "title",
    "date",
    "meeting date",
    "type",
    "meeting type",
    "category",
    "district",
    "summary",
    "objective",
    "outcome",
    "followUp",
    "followup",
    "followUpReason",
    "stakeholders",
    "companies",
    "keyDecisions",
    "discussionPoints",
    "nlpStats",
    "transcript",
    "deptId",
    "deptName",
    "department",
    "actualCost",
    "budgetUsed",
    "estimatedCost",
    "budgetNotes",
    "actions",
    "activityCategory",
    "activityId",
    "meetingID",
    "activityTitle",
    "role",
    "mainActivity",
    "linkPhoto",
    "linkPhotoUrl",
    "attach file",
    "activityType",
    "organizationType",
    "dateFrom",
    "dateTo",
    "representativePosition",
    "representativeName",
    "representativeDepartment",
    "activityObjective",
    "invitationfrom",
    "location meeting",
    "other reps",
    "recaps",
    "sltdepartment",
    "sltposition",
    "sltreps",
    "stfemail",
    "supemail",
    "updated by",
]

DEPARTMENT_SHEET_COLUMNS = ["id", "name", "budget"]
HISTORY_SHEET_COLUMNS = [
    "id",
    "user_id",
    "thread_key",
    "thread_date",
    "thread_title",
    "timestamp",
    "question",
    "answer",
    "meeting_id",
    "meeting_title",
    "context",
]

STATUSES = ["Pending", "In Progress", "Done", "Overdue", "Cancelled"]
MTG_TYPES = ["Virtual", "Physical", "Not Provided"]
CATEGORIES = ["External Meeting", "Internal Meeting", "Workshop"]
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
ACTIVITY_TYPE_OPTIONS = ["None", "Virtual", "Physical", "Both"]
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
