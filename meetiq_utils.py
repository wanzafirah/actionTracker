import calendar
import html
import json
import re
from datetime import date, datetime
from uuid import uuid4

import pandas as pd
import streamlit as st


def uid() -> str:
    return uuid4().hex[:9]


def today_str() -> str:
    return date.today().isoformat()


def generate_activity_id(category: str, meeting_date: date) -> str:
    category_clean = re.sub(r"[^A-Z]", "", category.upper())
    prefix = category_clean[:3] if category_clean else "ACT"
    return f"{prefix}-{meeting_date.strftime('%Y%m%d')}-{uuid4().hex[:4].upper()}"


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
    return [text for item in items or [] if (text := entity_text(item))]


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


def action_belongs_to_talentcorp(action: dict) -> bool:
    company = normalize_value(action.get("company"), "").strip().lower()
    owner = normalize_value(action.get("owner"), "").strip().lower()
    text = normalize_value(action.get("text"), "").strip().lower()
    allowed = ("", "none", "not stated", "internal", "talentcorp", "talent corp")
    if company in allowed or "talentcorp" in company or "talent corp" in company:
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


def render_plotly_chart(fig, use_container_width: bool = True):
    st.plotly_chart(fig, use_container_width=use_container_width, config={"displayModeBar": False})


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


def append_document_to_transcript(current_text: str, extracted_text: str) -> str:
    current_text = current_text.strip()
    if current_text:
        return f"{current_text}\n\nSupporting document:\n{extracted_text}"
    return extracted_text


def compact_transcript_for_prompt(text: str, max_chars: int = 800) -> str:
    text = re.sub(r"\s+", " ", text or "").strip()
    if len(text) <= max_chars:
        return text
    priority_lines = []
    chunks = re.split(r"(?<=[.!?])\s+", text)
    keywords = (
        "action", "deadline", "follow-up", "follow up", "decision", "agreed", "approved",
        "pending", "date", "month", "launch", "budget", "owner", "assign", "task",
    )
    for chunk in chunks:
        if any(keyword in chunk.lower() for keyword in keywords):
            priority_lines.append(chunk)
    compacted = " ".join(priority_lines)[:max_chars].strip()
    return compacted or text[:max_chars].strip()


def transcript_sentences(text: str) -> list:
    return [part.strip() for part in re.split(r"(?<=[.!?])\s+", text or "") if part.strip()]


def fallback_discussion_points(text: str, limit: int = 4) -> list:
    points = []
    for sentence in transcript_sentences(text):
        lowered = sentence.lower()
        if any(word in lowered for word in ("discuss", "review", "align", "explore", "present", "topic", "outline", "outcome")):
            points.append(sentence)
        if len(points) >= limit:
            break
    return points


def fallback_key_decisions(text: str, limit: int = 3) -> list:
    decisions = []
    for sentence in transcript_sentences(text):
        lowered = sentence.lower()
        if any(word in lowered for word in ("decided", "agreed", "approved", "confirmed", "will proceed", "was decided")):
            decisions.append(sentence)
        if len(decisions) >= limit:
            break
    return decisions


def fallback_action_items(text: str, limit: int = 5) -> list:
    actions = []
    seen = set()
    for sentence in transcript_sentences(text):
        lowered = sentence.lower()
        if "requested support for" in lowered:
            detail = sentence.split("requested support for", 1)[1].strip(" .")
            if detail:
                action_text = f"Provide support for {detail}"
                if action_text.lower() not in seen:
                    actions.append(
                        {
                            "text": action_text,
                            "owner": "TalentCorp team",
                            "company": "TalentCorp",
                            "deadline": "None",
                            "priority": "Medium",
                            "follow_up_required": True,
                            "follow_up_reason": "Support was explicitly requested during the meeting.",
                            "suggestion": f"Plan the next steps for {detail}.",
                            "ner_entities": [],
                        }
                    )
                    seen.add(action_text.lower())
        if "requested" in lowered and "briefing session" in lowered:
            action_text = "Prepare and coordinate ambassador programme briefing sessions"
            if action_text.lower() not in seen:
                actions.append(
                    {
                        "text": action_text,
                        "owner": "TalentCorp team",
                        "company": "TalentCorp",
                        "deadline": "None",
                        "priority": "Medium",
                        "follow_up_required": True,
                        "follow_up_reason": "Briefing support was explicitly requested.",
                        "suggestion": "Coordinate the briefing scope and schedule with the university.",
                        "ner_entities": [],
                    }
                )
                seen.add(action_text.lower())
        if "agreed to continue refining" in lowered:
            detail = sentence.split("agreed to continue refining", 1)[1].strip(" .")
            action_text = f"Continue refining {detail}" if detail else "Continue refining the programme framework"
            if action_text.lower() not in seen:
                actions.append(
                    {
                        "text": action_text,
                        "owner": "TalentCorp team",
                        "company": "TalentCorp",
                        "deadline": "None",
                        "priority": "Medium",
                        "follow_up_required": True,
                        "follow_up_reason": "The meeting ended with agreement to continue refinement.",
                        "suggestion": "Prepare an updated framework before the next discussion.",
                        "ner_entities": [],
                    }
                )
                seen.add(action_text.lower())
        if "requested" in lowered and "onboarding" in lowered:
            action_text = "Support student account onboarding for the programme"
            if action_text.lower() not in seen:
                actions.append(
                    {
                        "text": action_text,
                        "owner": "TalentCorp team",
                        "company": "TalentCorp",
                        "deadline": "None",
                        "priority": "Medium",
                        "follow_up_required": True,
                        "follow_up_reason": "Onboarding support was explicitly requested.",
                        "suggestion": "Define the onboarding process and coordination plan.",
                        "ner_entities": [],
                    }
                )
                seen.add(action_text.lower())
        if len(actions) >= limit:
            break
    return actions[:limit]


def is_objective_only_transcript(text: str) -> bool:
    lowered = (text or "").lower()
    objective_markers = ("purpose of this meeting", "meeting to discuss", "aiming to", "objective", "to align", "to discuss")
    action_markers = ("will do", "needs to", "must", "assigned", "action item", "follow up", "send", "prepare", "submit", "review by")
    return any(marker in lowered for marker in objective_markers) and not any(marker in lowered for marker in action_markers)


def build_meeting_dataframe(meetings: list) -> pd.DataFrame:
    if not meetings:
        return pd.DataFrame()
    rows = []
    for meeting in meetings:
        nlp_stats = meeting.get("nlpStats", {})
        parsed_date = pd.to_datetime(meeting.get("date"), errors="coerce")
        rows.append(
            {
                "id": meeting["id"],
                "title": meeting["title"],
                "date": meeting["date"],
                "month": parsed_date.strftime("%Y-%m") if not pd.isna(parsed_date) else "",
                "year": int(parsed_date.year) if not pd.isna(parsed_date) else None,
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


def get_upcoming_meetings(meetings: list, limit: int = 4, sort_order: str = "Earliest deadline") -> list:
    candidates = []
    for meeting in meetings:
        actions = meeting.get("actions", [])
        active_actions = [action for action in actions if normalize_status(action) in {"Pending", "In Progress"}]
        if not active_actions:
            continue

        deadlines = []
        for action in active_actions:
            deadline = normalize_value(action.get("deadline"), "")
            if not deadline or deadline == "None":
                continue
            try:
                deadlines.append(datetime.strptime(str(deadline), "%Y-%m-%d").date())
            except Exception:
                continue

        if deadlines:
            deadline_key = min(deadlines)
        else:
            meeting_date_text = normalize_value(meeting.get("date"), "")
            try:
                deadline_key = datetime.strptime(str(meeting_date_text), "%Y-%m-%d").date()
            except Exception:
                deadline_key = date.max if sort_order == "Earliest deadline" else date.min

        candidates.append((deadline_key, meeting))

    reverse = sort_order == "Latest deadline"
    ordered = sorted(candidates, key=lambda item: item[0], reverse=reverse)
    return [meeting for _, meeting in ordered[:limit]]


def get_pending_deadline_days(meetings: list, year: int, month: int) -> set:
    pending_days = set()
    for meeting in meetings:
        for action in meeting.get("actions", []):
            if normalize_status(action) != "Pending":
                continue
            deadline = normalize_value(action.get("deadline"), "")
            if not deadline or deadline == "None":
                continue
            try:
                parsed = datetime.strptime(str(deadline), "%Y-%m-%d").date()
            except Exception:
                continue
            if parsed.year == year and parsed.month == month:
                pending_days.add(parsed.day)
    return pending_days


def get_meetings_for_deadline(meetings: list, target_date: date) -> list:
    matched = []
    target_text = target_date.isoformat()
    for meeting in meetings:
        actions = meeting.get("actions", [])
        if not actions:
            continue
        if any(
            normalize_status(action) in {"Pending", "In Progress"} and normalize_value(action.get("deadline"), "") == target_text
            for action in actions
        ):
            matched.append(meeting)
    return matched


def build_calendar_html(meetings: list, year: int, month: int) -> str:
    cal = calendar.Calendar(firstweekday=0)
    weeks = cal.monthdayscalendar(year, month)
    day_labels = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
    header = "".join(f"<div class='calendar-day-label'>{label}</div>" for label in day_labels)
    cells = []
    today = date.today()
    pending_days = get_pending_deadline_days(meetings, year, month)
    for week in weeks:
        for day_num in week:
            classes = ["calendar-day"]
            label = "" if day_num == 0 else str(day_num)
            if day_num == 0:
                classes.append("empty")
            else:
                if day_num in pending_days:
                    classes.append("pending-deadline")
                if day_num == today.day and month == today.month and year == today.year:
                    classes.append("today")
            cells.append(f"<div class='{' '.join(classes)}'>{label}</div>")
    return (
        '<div class="calendar-widget">'
        f'<div class="calendar-grid calendar-head">{header}</div>'
        f'<div class="calendar-grid">{"".join(cells)}</div>'
        "</div>"
    )


def render_chat_bubble_html(role: str, text: str) -> str:
    safe_text = html.escape(str(text)).replace("\n", "<br>")
    return f'<div class="chat-bubble {role}">{safe_text}</div>'
