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


def meeting_id_prefix(category: str) -> str:
    category_text = str(category or "").strip().lower()
    mapping = {
        "external meeting": "EX",
        "internal meeting": "IN",
        "workshop": "WS",
    }
    if category_text in mapping:
        return mapping[category_text]
    cleaned = re.sub(r"[^A-Z]", "", str(category or "").upper())
    return cleaned[:2] if cleaned else "MT"


def generate_activity_id(category: str, meeting_date: date, meetings: list | None = None) -> str:
    prefix = meeting_id_prefix(category)
    year_text = str(meeting_date.year)
    next_number = 1
    if meetings:
        pattern = re.compile(rf"^{re.escape(prefix)}-{re.escape(year_text)}-(\d{{3}})$")
        existing_numbers = []
        for meeting in meetings:
            candidate = str(
                meeting.get("meetingID")
                or meeting.get("activityId")
                or meeting.get("id")
                or ""
            ).strip()
            matched = pattern.match(candidate)
            if matched:
                try:
                    existing_numbers.append(int(matched.group(1)))
                except Exception:
                    continue
        if existing_numbers:
            next_number = max(existing_numbers) + 1
    return f"{prefix}-{year_text}-{next_number:03d}"


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
    department = normalize_value(action.get("department") or action.get("company"), "").strip().lower()
    owner = normalize_value(action.get("owner"), "").strip().lower()
    text = normalize_value(action.get("text"), "").strip().lower()
    allowed = ("", "none", "not stated", "internal", "talentcorp", "talent corp")
    if department in allowed or "talentcorp" in department or "talent corp" in department:
        return True
    if any(token in department for token in ("external", "partner", "company", "university", "college", "school")):
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


def smart_summary_sentences(text: str, limit: int = 3) -> list[str]:
    sentences = transcript_sentences(text)
    if not sentences:
        return []

    keywords = (
        "objective",
        "purpose",
        "meeting",
        "discuss",
        "review",
        "align",
        "decision",
        "decided",
        "agreed",
        "confirmed",
        "action",
        "follow up",
        "follow-up",
        "deadline",
        "next step",
        "next steps",
        "summary",
        "outcome",
        "proposal",
        "plan",
        "schedule",
        "provided",
        "requested",
        "will",
    )
    scored: list[tuple[int, int, str]] = []
    for index, sentence in enumerate(sentences):
        lowered = sentence.lower()
        score = 0
        score += sum(1 for keyword in keywords if keyword in lowered)
        if len(sentence) > 35:
            score += 1
        if len(sentence) > 220:
            score -= 1
        if index == 0 and any(keyword in lowered for keyword in ("objective", "purpose", "discuss", "review", "decision", "request", "action", "follow up")):
            score += 1
        if index == len(sentences) - 1:
            score += 1
        scored.append((score, index, sentence))

    selected: list[str] = []
    seen = set()
    for _, _, sentence in sorted(scored, key=lambda item: (item[0], -item[1]), reverse=True):
        normalized = re.sub(r"\s+", " ", sentence).strip()
        key = normalized.lower()
        if not normalized or key in seen:
            continue
        selected.append(normalized)
        seen.add(key)
        if len(selected) >= limit:
            break

    selected.sort(key=lambda item: sentences.index(item) if item in sentences else 0)
    return selected


def fallback_discussion_points(text: str, limit: int = 4) -> list:
    points = []
    for sentence in transcript_sentences(text):
        lowered = sentence.lower()
        if any(word in lowered for word in ("discuss", "review", "align", "explore", "present", "topic", "outline", "outcome")):
            points.append(sentence)
        if len(points) >= limit:
            break
    if not points:
        points = smart_summary_sentences(text, limit=limit)
    return points


def fallback_key_decisions(text: str, limit: int = 3) -> list:
    decisions = []
    for sentence in transcript_sentences(text):
        lowered = sentence.lower()
        if any(word in lowered for word in ("decided", "agreed", "approved", "confirmed", "will proceed", "was decided")):
            decisions.append(sentence)
        if len(decisions) >= limit:
            break
    if not decisions:
        sentences = smart_summary_sentences(text, limit=limit)
        decisions = [sentence for sentence in sentences if any(word in sentence.lower() for word in ("decided", "agreed", "confirmed", "will", "next step", "action"))]
    return decisions


def smart_summary_from_transcript(text: str, limit: int = 5) -> str:
    sentences = smart_summary_sentences(text, limit=limit)
    if sentences:
        return " ".join(sentences).strip()
    compact = re.sub(r"\s+", " ", text or "").strip()
    return compact[:420].rsplit(" ", 1)[0].rstrip(".,;:") + "..." if len(compact) > 420 else compact


def normalize_compare_text(text: str) -> str:
    return re.sub(r"\s+", " ", re.sub(r"[^\w\s]", "", text or "").lower()).strip()


def first_meaningful_sentence(text: str) -> str:
    for sentence in transcript_sentences(text):
        if len(sentence.split()) >= 6:
            return sentence
    sentences = transcript_sentences(text)
    return sentences[0] if sentences else ""


def looks_like_copied_intro(candidate: str, transcript: str) -> bool:
    candidate_norm = normalize_compare_text(candidate)
    if not candidate_norm:
        return True
    if len(candidate_norm.split()) < 8:
        return True

    intro = normalize_compare_text(first_meaningful_sentence(transcript))
    if not intro:
        return False

    if candidate_norm == intro or candidate_norm in intro or intro in candidate_norm:
        return True

    candidate_words = candidate_norm.split()
    intro_words = intro.split()
    overlap = len(set(candidate_words) & set(intro_words))
    if intro_words and overlap / max(len(set(intro_words)), 1) >= 0.8 and len(candidate_words) <= len(intro_words) + 4:
        return True

    return False


def summary_needs_expansion(candidate: str, transcript: str, min_words: int = 45) -> bool:
    candidate_text = normalize_compare_text(candidate)
    if not candidate_text:
        return True
    if len(candidate_text.split()) < min_words:
        return True
    if looks_like_copied_intro(candidate, transcript):
        return True
    return False


def better_objective_from_transcript(text: str) -> str:
    sentences = transcript_sentences(text)
    if not sentences:
        return ""

    keywords = (
        "objective",
        "purpose",
        "discuss",
        "review",
        "align",
        "brainstorm",
        "plan",
        "explore",
        "propose",
        "request",
        "next step",
        "follow up",
        "collaboration",
    )
    scored: list[tuple[int, int, str]] = []
    for index, sentence in enumerate(sentences):
        lowered = sentence.lower()
        score = sum(1 for keyword in keywords if keyword in lowered)
        if len(sentence.split()) >= 8:
            score += 1
        if index == 0:
            score -= 1
        scored.append((score, index, sentence))

    chosen = max(scored, key=lambda item: (item[0], -item[1]))[2]
    cleaned = re.sub(
        r"^(today|this morning|this afternoon|yesterday|the team|the meeting|our team|we)\b[,:-]?\s*",
        "",
        chosen,
        flags=re.IGNORECASE,
    ).strip()
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    if len(cleaned.split()) > 34:
        cleaned = " ".join(cleaned.split()[:34]).rstrip(".,;:") + "..."
    return cleaned


def extract_deadline_phrase(text: str) -> str:
    lowered = normalize_value(text, "").lower()
    patterns = [
        r"\bby\s+([A-Z][a-z]+\s+\d{1,2},?\s+\d{4})",
        r"\bby\s+([A-Z][a-z]+\s+\d{4})",
        r"\bby\s+(early\s+[A-Z][a-z]+)",
        r"\bby\s+(mid\s+[A-Z][a-z]+)",
        r"\bby\s+(late\s+[A-Z][a-z]+)",
        r"\b(in\s+early\s+[A-Z][a-z]+)",
        r"\b(in\s+mid\s+[A-Z][a-z]+)",
        r"\b(in\s+late\s+[A-Z][a-z]+)",
        r"\b(in\s+[A-Z][a-z]+)",
        r"\b(before\s+the\s+programme\s+begins)",
        r"\b(before\s+the\s+program(?:me)?\s+begins)",
        r"\b(before\s+the\s+session)",
        r"\b(before\s+the\s+workshop)",
        r"\b(before\s+the\s+event)",
    ]
    source = normalize_value(text, "")
    for pattern in patterns:
        match = re.search(pattern, source, flags=re.IGNORECASE)
        if match:
            return normalize_value(match.group(1), "None")
    if "early october" in lowered:
        return "Early October"
    if "before the programme begins" in lowered or "before the program begins" in lowered:
        return "Before the programme begins"
    return "None"


def fallback_action_items(text: str, limit: int = 5) -> list:
    actions = []
    seen = set()
    text_lower = (text or "").lower()

    def add_action(action_text: str, owner: str, department: str, suggestion: str, reason: str, deadline: str = "None"):
        if action_text.lower() in seen or len(actions) >= limit:
            return
        actions.append(
            {
                "text": action_text,
                "owner": owner,
                "department": department,
                "deadline": deadline,
                "priority": "Medium",
                "follow_up_required": True,
                "follow_up_reason": reason,
                "suggestion": suggestion,
                "ner_entities": [],
            }
        )
        seen.add(action_text.lower())

    for sentence in transcript_sentences(text):
        lowered = sentence.lower()
        deadline_text = extract_deadline_phrase(sentence)
        if "requested support for" in lowered:
            detail = sentence.split("requested support for", 1)[1].strip(" .")
            if detail:
                add_action(
                    f"Follow up on support requested for {detail}",
                    "TalentCorp team",
                    "TalentCorp",
                    f"Clarify the required support and confirm the next steps for {detail}.",
                    "Support was explicitly requested during the meeting.",
                    deadline=deadline_text,
                )
        if "requested" in lowered and "briefing session" in lowered:
            add_action(
                "Prepare and coordinate briefing sessions",
                "TalentCorp team",
                "TalentCorp",
                "Confirm the briefing scope, timing, and participants before the session.",
                "Briefing support was explicitly requested.",
                deadline=deadline_text,
            )
        if "agreed to continue refining" in lowered:
            detail = sentence.split("agreed to continue refining", 1)[1].strip(" .")
            add_action(
                f"Continue refining {detail}" if detail else "Continue refining the programme framework",
                "TalentCorp team",
                "TalentCorp",
                "Prepare an updated framework before the next discussion.",
                "The meeting ended with agreement to continue refinement.",
                deadline=deadline_text,
            )
        if "requested" in lowered and "onboarding" in lowered:
            add_action(
                "Support student account onboarding for the programme",
                "TalentCorp team",
                "TalentCorp",
                "Define the onboarding process and coordination plan.",
                "Onboarding support was explicitly requested.",
                deadline=deadline_text,
            )
        if "talentcorp agreed to support" in lowered or "talentcorp will support" in lowered:
            detail = sentence
            detail = re.sub(r"^.*?(?:agreed to support|will support)\s+", "", detail, flags=re.IGNORECASE).strip(" .")
            add_action(
                f"Support {detail}" if detail else "Support the planned session",
                "TalentCorp team",
                "TalentCorp",
                "Coordinate the required speakers, materials, and session support.",
                "TalentCorp support was explicitly confirmed during the meeting.",
                deadline=deadline_text,
            )
        if "providing speakers" in lowered or "relevant materials" in lowered:
            add_action(
                "Provide speakers and relevant materials for the session",
                "TalentCorp team",
                "TalentCorp",
                "Confirm the speaker lineup and prepare the supporting materials before the session.",
                "TalentCorp committed to support the session with speakers and materials.",
                deadline=deadline_text,
            )
        if "will circulate" in lowered and any(token in lowered for token in ("participant list", "logistical details", "logistics")):
            add_action(
                "Follow up with the university team on the updated participant list and logistical details",
                "TalentCorp team",
                "TalentCorp",
                "Track the updated participant list and logistics details before the programme starts.",
                "The university team committed to circulate updated participant and logistics information.",
                deadline=deadline_text,
            )
        if "will provide" in lowered and "talentcorp" in lowered:
            detail = re.sub(r"^.*?will provide\s+", "", sentence, flags=re.IGNORECASE).strip(" .")
            add_action(
                f"Provide {detail}" if detail else "Provide the requested deliverables",
                "TalentCorp team",
                "TalentCorp",
                "Confirm the deliverables and prepare them for the next step.",
                "TalentCorp committed to provide a deliverable during the meeting.",
                deadline=deadline_text,
            )
    if any(phrase in text_lower for phrase in ("position name", "salary range", "required skills", "number of openings", "upskilling opportunities")):
        add_action(
            "Follow up with WD to provide required program details",
            "TalentCorp team",
            "TalentCorp",
            "Request the missing details and confirm the submission timeline with WD.",
            "The recap explicitly says WD needs to provide program information before the initiative can proceed.",
        )
    if "mywira" in text_lower or "program kesinambungan kerjaya veteran" in text_lower:
        add_action(
            "Coordinate WD onboarding for the MyWira initiative",
            "TalentCorp team",
            "TalentCorp",
            "Confirm WD participation and align the next steps for the programme.",
            "The recap says WD is interested in joining the programme and TalentCorp agreed to the collaboration.",
        )
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
                    "department": normalize_value(action.get("department") or action.get("company"), "Not stated"),
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
