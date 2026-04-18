import asyncio
import json
import os
import re
import tempfile
from datetime import datetime, timezone
from io import BytesIO

import requests
from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import JSONResponse, PlainTextResponse

from meetiq_constants import HISTORY_SHEET_NAME, MEETINGS_SHEET_NAME
from meetiq_services import call_ollama, extract_text_from_document, transcribe_audio_file
from meetiq_utils import (
    fallback_action_items,
    fallback_discussion_points,
    fallback_key_decisions,
    compact_transcript_for_prompt,
    join_list,
    json_dumps_safe,
    normalize_status,
    normalize_value,
    is_objective_only_transcript,
    smart_summary_from_transcript,
    today_str,
    uid,
)


app = FastAPI(title="MeetIQ Telegram Backend", version="1.0.0")
_RECENT_UPDATE_IDS: set[int] = set()
_RECENT_MESSAGE_KEYS: set[str] = set()
_RECENT_CACHE_LIMIT = 500
NO_REPLY = "__NO_REPLY__"
_CHAT_SESSION_VERSION: dict[str, int] = {}


def get_telegram_token() -> str:
    return os.getenv("TELEGRAM_BOT_TOKEN", "").strip()


def get_telegram_webhook_secret() -> str:
    return os.getenv("TELEGRAM_WEBHOOK_SECRET", "").strip()


def get_supabase_config() -> dict:
    url = os.getenv("SUPABASE_URL", "").rstrip("/")
    key = (
        os.getenv("SUPABASE_SERVICE_ROLE_KEY", "")
        or os.getenv("SUPABASE_KEY", "")
        or os.getenv("SUPABASE_ANON_KEY", "")
    ).strip()
    return {"url": url, "key": key}


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


def supabase_select(table_name: str, params: dict) -> list[dict]:
    headers = supabase_headers()
    table_url = supabase_table_url(table_name)
    if headers is None or table_url is None:
        raise RuntimeError("Supabase storage is not configured.")

    query = {"select": "*"}
    query.update(params)
    response = requests.get(table_url, headers=headers, params=query, timeout=30)
    response.raise_for_status()
    payload = response.json()
    return payload if isinstance(payload, list) else []


def supabase_upsert(table_name: str, rows: list[dict]) -> None:
    if not rows:
        return

    headers = supabase_headers(prefer="resolution=merge-duplicates,return=minimal")
    table_url = supabase_table_url(table_name)
    if headers is None or table_url is None:
        raise RuntimeError("Supabase storage is not configured.")

    response = requests.post(
        table_url,
        headers=headers,
        params={"on_conflict": "id"},
        json=rows,
        timeout=30,
    )
    response.raise_for_status()


def supabase_insert(table_name: str, rows: list[dict]) -> None:
    if not rows:
        return

    headers = supabase_headers(prefer="return=minimal")
    table_url = supabase_table_url(table_name)
    if headers is None or table_url is None:
        raise RuntimeError("Supabase storage is not configured.")

    response = requests.post(table_url, headers=headers, json=rows, timeout=30)
    response.raise_for_status()


def get_telegram_api_base() -> str:
    token = get_telegram_token()
    if not token:
        raise RuntimeError("TELEGRAM_BOT_TOKEN is missing.")
    return f"https://api.telegram.org/bot{token}"


def telegram_api(method: str, payload: dict | None = None, files=None, timeout: int = 30) -> dict:
    url = f"{get_telegram_api_base()}/{method}"
    response = requests.post(url, json=payload or {}, files=files, timeout=timeout)
    response.raise_for_status()
    data = response.json()
    if not data.get("ok", False):
        raise RuntimeError(data.get("description", "Telegram request failed."))
    return data


def telegram_get_file(file_id: str) -> dict:
    url = f"{get_telegram_api_base()}/getFile"
    response = requests.get(url, params={"file_id": file_id}, timeout=30)
    response.raise_for_status()
    data = response.json()
    if not data.get("ok", False):
        raise RuntimeError(data.get("description", "Could not fetch Telegram file metadata."))
    return data["result"]


class TelegramUpload(BytesIO):
    def __init__(self, data: bytes, name: str):
        super().__init__(data)
        self.name = name


def download_telegram_file(file_id: str, file_name: str = "upload.bin") -> TelegramUpload:
    file_info = telegram_get_file(file_id)
    file_path = file_info.get("file_path")
    if not file_path:
        raise RuntimeError("Telegram file path is missing.")

    url = f"https://api.telegram.org/file/bot{get_telegram_token()}/{file_path}"
    response = requests.get(url, timeout=120)
    response.raise_for_status()
    return TelegramUpload(response.content, file_name)


def normalize_chat_user_id(value) -> str:
    return normalize_value(value, "").strip()


def parse_lines(text: str) -> list[str]:
    return [line.strip() for line in str(text or "").splitlines() if line.strip()]


def dedupe_preserve_order(items: list[str]) -> list[str]:
    seen = set()
    cleaned = []
    for item in items or []:
        value = normalize_value(item, "").strip()
        key = value.lower()
        if value and key not in seen:
            seen.add(key)
            cleaned.append(value)
    return cleaned


def has_action_signals(text: str) -> bool:
    lowered = (text or "").lower()
    return any(
        marker in lowered
        for marker in [
            "action item",
            "action items",
            "follow up",
            "follow-up",
            "next step",
            "next steps",
            "needs to",
            "need to",
            "should",
            "must",
            "please",
            "to proceed",
            "deadline",
            "assign",
            "assigned",
            "will start",
            "will send",
            "will provide",
        ]
    )


def split_sentences(text: str) -> list[str]:
    cleaned = re.sub(r"\s+", " ", str(text or "").strip())
    if not cleaned:
        return []
    parts = re.split(r"(?<=[.!?])\s+", cleaned)
    return [part.strip() for part in parts if part.strip()]


def fallback_title_from_text(raw_text: str) -> str:
    first_line = next((line.strip() for line in str(raw_text or "").splitlines() if line.strip()), "")
    if first_line:
        if len(first_line) <= 90:
            return first_line[:90]
        words = first_line.split()
        return " ".join(words[:8]).strip() or "Telegram meeting recap"
    return "Telegram meeting recap"


def normalize_title_key(value: str) -> str:
    return re.sub(r"\s+", " ", normalize_value(value, "").strip().lower())


def fallback_summary_from_text(raw_text: str, limit: int = 420) -> str:
    summary = smart_summary_from_transcript(raw_text, limit=3)
    if not summary:
        return "Meeting recap submitted via Telegram."
    if len(summary) > limit:
        return summary[:limit].rsplit(" ", 1)[0].rstrip(".,;:") + "..."
    return summary


def extract_explicit_action_items_from_text(raw_text: str) -> list[dict]:
    lines = [line.strip() for line in str(raw_text or "").splitlines()]
    action_start = None
    for index, line in enumerate(lines):
        lowered = line.lower()
        if lowered.startswith("action items") or lowered.startswith("action item") or lowered.startswith("next steps"):
            action_start = index + 1
            break

    if action_start is None:
        return []

    collected = []
    for line in lines[action_start:]:
        lowered = line.lower().strip()
        if not lowered:
            if collected:
                break
            continue
        if lowered.startswith(("summary:", "objective:", "outcome:", "stakeholders:", "meeting:", "recap:")):
            break
        if lowered.startswith(("-", "*", "•")) or re.match(r"^\d+[\.\)]\s+", lowered):
            item = re.sub(r"^[-*•]\s*", "", line).strip()
            item = re.sub(r"^\d+[\.\)]\s*", "", item).strip()
            if item:
                collected.append(
                    {
                        "text": item,
                        "owner": "Not stated",
                        "department": "Not stated",
                        "deadline": "None",
                        "status": "Pending",
                        "suggestion": "",
                        "priority": "Medium",
                        "follow_up_required": True,
                        "follow_up_reason": "Explicit action item listed in the submission.",
                        "ner_entities": [],
                    }
                )
        elif collected and not line.startswith((" ", "\t")):
            break

    return collected


def extract_meeting_title_from_question(question: str) -> str:
    text = normalize_value(question, "").strip()
    if not text:
        return ""

    patterns = [
        r"(?:title|meeting title|recap title)\s*[:\-]\s*(.+)$",
        r"(?:for|about|regarding|on)\s+(.+?)(?:\?|$)",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            candidate = normalize_value(match.group(1), "")
            candidate = re.sub(r"\b(recaps?|summary|action items?)\b.*$", "", candidate, flags=re.IGNORECASE).strip()
            return candidate

    quoted = re.search(r"\"([^\"]+)\"|'([^']+)'", text)
    if quoted:
        return normalize_value(quoted.group(1) or quoted.group(2), "")
    return ""


def is_meeting_lookup_question(question: str) -> bool:
    lowered = normalize_value(question, "").lower().strip()
    if not lowered:
        return False
    return bool(
        re.match(r"^(do we have|do we have any|is there|are there|have we|did we have|when is|when was)\b", lowered)
        or "meeting with" in lowered
        or "meet with" in lowered
        or "meeting for" in lowered
        or "meeting about" in lowered
    )


def build_lookup_response(question: str, meetings: list[dict]) -> tuple[str, dict | None]:
    tokens = [token for token in re.findall(r"[a-zA-Z0-9&]+", question.lower()) if len(token) >= 3]
    scored = []
    for meeting in meetings:
        score = score_meeting(question, meeting)
        if score > 0:
            scored.append((score, meeting))

    matched_meetings = [meeting for _, meeting in sorted(scored, key=lambda item: item[0], reverse=True)]
    if not matched_meetings:
        return "No matching meeting found in the saved data.", None

    top = matched_meetings[0]
    title = normalize_value(top.get("title"), "Untitled meeting")
    date_text = normalize_value(top.get("date"), "No date")
    company_text = join_list(top.get("companies", []), "None")
    stakeholder_text = join_list(top.get("stakeholders", []), "None")
    action_text = f"{len(top.get('actions', []) or [])} action item(s)"
    lead = tokens[-1] if tokens else "the requested meeting"

    lines = [
        f"Yes, I found {len(matched_meetings)} matching meeting(s) for {lead}.",
        f"Top match: {title} | {date_text}",
        f"Companies: {company_text}",
        f"Stakeholders: {stakeholder_text}",
        f"Action items: {action_text}",
    ]

    if len(matched_meetings) > 1:
        lines.append("Other matches:")
        for meeting in matched_meetings[1:4]:
            lines.append(
                f"- {normalize_value(meeting.get('title'), 'Untitled meeting')} | "
                f"{normalize_value(meeting.get('date'), 'No date')}"
            )

    return "\n".join(lines), top


def extract_people_from_text(raw_text: str) -> list[str]:
    intro_patterns = [
        r"^([A-Z][a-z]+(?:,\s*[A-Z][a-z]+)*(?:\s+and\s+[A-Z][a-z]+)?)\s+had a virtual meeting with\b",
        r"^([A-Z][a-z]+(?:\s*,\s*[A-Z][a-z]+)*(?:\s+and\s+[A-Z][a-z]+)?)\s+met with\b",
    ]
    patterns = [
        r"\b(?:given by|by|with|from|presented by|shared by|led by)\s+([A-Z][a-z]+(?:\s+[A-Z][a-z]+){0,3})",
        r"\b([A-Z][a-z]+(?:\s+[A-Z][a-z]+){0,3})\s+(?:explained|presented|shared|mentioned|said|noted)",
    ]
    candidates: list[str] = []
    for pattern in intro_patterns:
        match = re.search(pattern, raw_text or "", flags=re.IGNORECASE)
        if match:
            people_blob = match.group(1)
            people_blob = people_blob.replace(" and ", ",")
            for name in re.split(r"\s*,\s*", people_blob):
                value = normalize_value(name, "")
                if value:
                    candidates.append(value)
    for pattern in patterns:
        for match in re.findall(pattern, raw_text or ""):
            value = normalize_value(match, "")
            if value:
                candidates.append(value)
    return dedupe_preserve_order(candidates)


def extract_organization_candidates(raw_text: str) -> list[str]:
    organization_keywords = (
        "sdn bhd",
        "berhad",
        "corp",
        "corporation",
        "company",
        "ministry",
        "authority",
        "department",
        "agency",
        "institute",
        "university",
        "college",
        "board",
        "group",
        "association",
        "foundation",
        "council",
        "centre",
        "center",
        "office",
        "team",
    )
    patterns = [
        r"\b([A-Z][A-Za-z0-9&'./-]*(?:\s+[A-Z][A-Za-z0-9&'./-]*){0,5}\s+(?:Sdn Bhd|Berhad|Corp|Corporation|Company|Ministry|Authority|Department|Agency|Institute|University|College|Board|Group|Association|Foundation|Council|Centre|Center|Office|Team))\b",
        r"\b(?:of|with|from|at)\s+([A-Z][A-Za-z0-9&'./-]*(?:\s+[A-Z][A-Za-z0-9&'./-]*){0,5})",
    ]
    candidates: list[str] = []
    lowered = (raw_text or "").lower()
    for pattern in patterns:
        for match in re.findall(pattern, raw_text or ""):
            value = normalize_value(match, "")
            if not value:
                continue
            if any(keyword in value.lower() for keyword in organization_keywords) or value.lower() in lowered:
                candidates.append(value)
    return dedupe_preserve_order(candidates)


def extract_json_blob(raw: str) -> dict:
    if not raw:
        return {}

    text = raw.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?", "", text.strip(), flags=re.IGNORECASE).strip()
        text = re.sub(r"```$", "", text.strip()).strip()

    first = text.find("{")
    last = text.rfind("}")
    if first < 0 or last <= first:
        return {}

    candidate = text[first : last + 1]
    try:
        parsed = json.loads(candidate)
        return parsed if isinstance(parsed, dict) else {}
    except Exception:
        return {}


def repair_json_with_ollama(raw: str) -> dict:
    repair_system = "You repair malformed JSON. Return only valid JSON with no markdown."
    repair_prompt = f"Fix this into valid JSON and keep the same structure:\n\n{raw}"
    repaired = call_ollama(repair_system, repair_prompt, max_tokens=300)
    return extract_json_blob(repaired)


def chat_thread_key(user_id: str, meeting_date: str, meeting_title: str, meeting_id: str) -> str:
    return "::".join(
        [
            normalize_chat_user_id(user_id) or "anonymous",
            normalize_value(meeting_date, today_str()),
            normalize_value(meeting_title, "General"),
            normalize_value(meeting_id, ""),
        ]
    )


def meeting_context_id(meeting: dict) -> str:
    return normalize_value(meeting.get("id") or meeting.get("meetingID") or meeting.get("activityId"), "")


def meeting_context_text(meeting: dict) -> str:
    actions = meeting.get("actions", []) or []
    action_lines = []
    for action in actions:
        action_lines.append(
            f"- {normalize_value(action.get('text'))} | owner: {normalize_value(action.get('owner'), 'Not stated')} | "
            f"department: {normalize_value(action.get('department') or action.get('company'), 'Not stated')} | "
            f"deadline: {normalize_value(action.get('deadline'), 'None')} | status: {normalize_status(action)}"
        )

    return "\n".join(
        [
            f"Date: {normalize_value(meeting.get('date'), today_str())}",
            f"Title: {normalize_value(meeting.get('title'), 'Untitled meeting')}",
            f"Summary: {normalize_value(meeting.get('summary') or meeting.get('recaps'), 'No summary available.')}",
            f"Objective: {normalize_value(meeting.get('objective'), 'Not provided')}",
            f"Outcome: {normalize_value(meeting.get('outcome'), 'Not provided')}",
            f"Stakeholders: {join_list(meeting.get('stakeholders', []), 'None')}",
            f"Companies: {join_list(meeting.get('companies', []), 'None')}",
            "Action items:",
            "\n".join(action_lines) if action_lines else "- None",
        ]
    )


def build_history_entry(
    *,
    user_id: str,
    thread_key: str,
    thread_date: str,
    thread_title: str,
    question: str,
    answer: str,
    meeting_id: str = "",
    meeting_title: str = "",
    context: str = "General",
) -> dict:
    return {
        "id": uid(),
        "user_id": normalize_chat_user_id(user_id),
        "thread_key": thread_key,
        "thread_date": thread_date,
        "thread_title": thread_title,
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "question": question,
        "answer": answer,
        "meeting_id": meeting_id,
        "meeting_title": meeting_title,
        "context": context,
    }


def build_meeting_record(raw_text: str, recap: dict, user_id: str, source_name: str) -> dict:
    meeting_id = uid()
    meeting_title = normalize_value(recap.get("title"), "")
    if meeting_title in ("None", "Untitled meeting"):
        meeting_title = fallback_title_from_text(raw_text) if raw_text.strip() else normalize_value(source_name, "Untitled meeting")

    summary = normalize_value(recap.get("summary"), "")
    if summary in ("", "None", "No summary available."):
        summary = fallback_summary_from_text(raw_text)
    objective = normalize_value(recap.get("objective"), "")
    if objective in ("", "None", "Not provided"):
        objective = "Review the meeting discussion and align on the next steps."
    outcome = normalize_value(recap.get("outcome"), "")
    if outcome in ("", "None", "Not provided"):
        outcome = "Meeting recap captured from the submitted text."
    key_decisions = dedupe_preserve_order(recap.get("key_decisions", []) or fallback_key_decisions(raw_text))
    discussion_points = dedupe_preserve_order(recap.get("discussion_points", []) or fallback_discussion_points(raw_text))
    action_items = recap.get("action_items", []) or fallback_action_items(raw_text)
    people_from_text = extract_people_from_text(raw_text)
    organizations_from_text = extract_organization_candidates(raw_text)

    normalized_actions = []
    for index, action in enumerate(action_items):
        normalized_actions.append(
            {
                "id": f"{meeting_id}_a{index}",
                "text": normalize_value(action.get("text"), "Follow up required"),
                "owner": normalize_value(action.get("owner"), "Not stated"),
                "department": normalize_value(action.get("department") or action.get("company"), "Not stated"),
                "deadline": normalize_value(action.get("deadline"), "None"),
                "status": normalize_value(action.get("status"), "Pending"),
                "priority": normalize_value(action.get("priority"), "Medium"),
                "suggestion": normalize_value(action.get("suggestion"), ""),
                "follow_up_required": bool(action.get("follow_up_required", True)),
                "follow_up_reason": normalize_value(action.get("follow_up_reason"), ""),
                "ner_entities": action.get("ner_entities", []),
            }
        )

    if not discussion_points:
        discussion_points = ["Meeting recap generated from user submission."]

    return {
        "id": meeting_id,
        "user_id": normalize_chat_user_id(user_id),
        "title": meeting_title,
        "date": today_str(),
        "meeting date": today_str(),
        "type": normalize_value(recap.get("meeting_type"), "Not Provided"),
        "meeting type": normalize_value(recap.get("meeting_type"), "Not Provided"),
        "category": normalize_value(recap.get("category"), "Internal Meeting"),
        "district": normalize_value(recap.get("district"), ""),
        "summary": summary,
        "recaps": summary,
        "objective": objective,
        "outcome": outcome,
        "followUp": bool(normalized_actions),
        "followup": "Yes" if normalized_actions else "No",
        "followUpReason": normalize_value(recap.get("follow_up_reason"), ""),
        "stakeholders": dedupe_preserve_order((recap.get("stakeholders", []) or []) + people_from_text),
        "companies": dedupe_preserve_order((recap.get("companies", []) or []) + organizations_from_text),
        "keyDecisions": key_decisions,
        "discussionPoints": discussion_points,
        "nlpStats": recap.get("nlpStats", {}) or {},
        "transcript": raw_text,
        "deptId": "",
        "deptName": normalize_value(recap.get("department"), ""),
        "department": normalize_value(recap.get("department"), ""),
        "actualCost": 0,
        "budgetUsed": 0,
        "estimatedCost": 0,
        "budgetNotes": "",
        "actions": normalized_actions,
        "activityCategory": normalize_value(recap.get("activityCategory"), "External Meeting"),
        "activityId": meeting_id,
        "meetingID": meeting_id,
        "activityTitle": meeting_title,
        "role": normalize_value(recap.get("role"), ""),
        "mainActivity": "Yes",
        "linkPhoto": "",
        "linkPhotoUrl": "",
        "attach file": source_name,
        "activityType": normalize_value(recap.get("activityType"), "None"),
        "organizationType": normalize_value(recap.get("organizationType"), "Company"),
        "dateFrom": today_str(),
        "dateTo": today_str(),
        "representativePosition": "",
        "representativeName": "",
        "representativeDepartment": "",
        "activityObjective": objective,
        "invitationfrom": "",
        "location meeting": "",
        "other reps": "",
        "recaps": summary,
        "sltdepartment": normalize_value(recap.get("department"), ""),
        "sltposition": "",
        "sltreps": "",
        "stfemail": "",
        "supemail": "",
        "updated by": "Telegram Bot",
    }


def summarize_meeting_text(raw_text: str, source_name: str = "Telegram message") -> dict:
    prompt_text = compact_transcript_for_prompt(raw_text, max_chars=3000)
    objective_only = is_objective_only_transcript(prompt_text)
    has_explicit_actions = bool(extract_explicit_action_items_from_text(raw_text)) or has_action_signals(raw_text)
    system = """You are MeetIQ.
Turn meeting notes, long text, or transcript into a structured meeting recap.
Return valid JSON only. No markdown, no commentary.

Required JSON keys:
- title
- summary
- objective
- outcome
- follow_up
- follow_up_reason
- key_decisions (array of strings)
- discussion_points (array of strings)
- action_items (array of objects with text, owner, department, deadline, status, suggestion)
- stakeholders (array of strings)
- companies (array of strings)
- department
- organizationType

Rules:
- If the text is a meeting recap or minutes, always extract a useful summary even if the text is long.
- The summary should be 6 to 8 sentences when the discussion is long, and it should capture the main context, decision, concern, request, and next step.
- The objective should be a concise paraphrase of the meeting purpose, not a copy of the opening recap line.
- The summary should synthesize the full discussion in your own words; do not restate the first transcript sentence verbatim.
- Preserve the original meeting context and key participants. Avoid generic filler such as "hello everyone" unless it is truly relevant.
- The objective should reflect what the meeting was trying to achieve.
- The outcome should say what was decided, agreed, or what remains pending.
- Only create action items when they are explicitly written in the text or clearly stated as a task/request.
- Do not invent action items from general discussion.
- Only list people in stakeholders; do not put organizations or companies in stakeholders.
- Only list companies or organizations in companies.
- If the meeting has no clear task or request, return an empty action_items list.
- If the transcript is discussion-heavy, summarize the whole conversation instead of only the opening lines.
"""
    if objective_only:
        system += "\n- The source text may be objective-style meeting notes; infer the recap from the context and do not leave fields blank."
    if not has_explicit_actions:
        system += "\n- This submission may be discussion-heavy; prefer a clean summary and leave action_items empty unless the text clearly lists tasks."

    user_msg = f"Source name: {source_name}\nDate: {today_str()}\n\nText:\n{prompt_text}"
    raw = call_ollama(system, user_msg, max_tokens=1200)
    parsed = extract_json_blob(raw)
    if not parsed:
        parsed = repair_json_with_ollama(raw)

    parsed["title"] = normalize_value(parsed.get("title"), "") or fallback_title_from_text(raw_text)
    parsed["summary"] = normalize_value(parsed.get("summary"), "") or fallback_summary_from_text(raw_text)
    parsed["objective"] = normalize_value(parsed.get("objective"), "") or "Review the meeting discussion and align on the next steps."
    parsed["outcome"] = normalize_value(parsed.get("outcome"), "") or "Meeting recap captured from the submitted text."
    parsed["key_decisions"] = dedupe_preserve_order(parsed.get("key_decisions", []) or fallback_key_decisions(raw_text))
    parsed["discussion_points"] = dedupe_preserve_order(parsed.get("discussion_points", []) or fallback_discussion_points(raw_text))
    explicit_actions = extract_explicit_action_items_from_text(raw_text)
    action_items = explicit_actions if explicit_actions else ([] if not has_action_signals(raw_text) else fallback_action_items(raw_text))
    parsed["action_items"] = action_items
    parsed["stakeholders"] = dedupe_preserve_order((parsed.get("stakeholders", []) or []) + extract_people_from_text(raw_text))
    parsed["companies"] = dedupe_preserve_order((parsed.get("companies", []) or []) + extract_organization_candidates(raw_text))
    parsed["stakeholders"] = [item for item in parsed["stakeholders"] if not any(keyword in item.lower() for keyword in ("sdn bhd", "berhad", "corp", "company", "ministry", "authority", "department", "agency", "university", "college", "group", "team"))]
    parsed["companies"] = [item for item in parsed["companies"] if item.lower() not in {person.lower() for person in parsed["stakeholders"]}]
    parsed.setdefault("follow_up", bool(parsed.get("action_items")))
    parsed.setdefault("follow_up_reason", "")
    parsed.setdefault("department", "")
    parsed.setdefault("organizationType", "Company")
    parsed.setdefault("activityType", "None")
    parsed.setdefault("meeting_type", "Not Provided")
    parsed.setdefault("category", "Internal Meeting")
    parsed.setdefault("nlpStats", {})
    return parsed


def load_user_meetings(user_id: str, limit: int = 100) -> list[dict]:
    try:
        rows = supabase_select(
            MEETINGS_SHEET_NAME,
            {
                "user_id": f"eq.{normalize_chat_user_id(user_id)}",
                "order": "date.desc",
                "limit": str(limit),
            },
        )
        return rows
    except Exception:
        return []


def load_user_history(user_id: str, limit: int = 200) -> list[dict]:
    try:
        rows = supabase_select(
            HISTORY_SHEET_NAME,
            {
                "user_id": f"eq.{normalize_chat_user_id(user_id)}",
                "order": "timestamp.desc",
                "limit": str(limit),
            },
        )
        return rows
    except Exception:
        return []


def question_tokens(question: str) -> set[str]:
    lowered = question.lower().strip()
    return {
        token
        for token in re.findall(r"[a-zA-Z0-9&]+", lowered)
        if len(token) >= 3
        and token
        not in {
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
            "recap",
            "recaps",
            "summary",
            "summarize",
            "summarise",
            "details",
            "detail",
            "past",
            "previous",
            "solution",
            "solve",
            "help",
        }
    }


def score_meeting(question: str, meeting: dict, last_meeting_id: str = "") -> int:
    tokens = question_tokens(question)
    blob = meeting_context_text(meeting).lower()
    score = sum(1 for token in tokens if token in blob)
    title = normalize_value(meeting.get("title"), "").lower()
    companies_blob = join_list(meeting.get("companies", []), "").lower()
    stakeholders_blob = join_list(meeting.get("stakeholders", []), "").lower()
    if any(token in title or token in companies_blob or token in stakeholders_blob for token in tokens):
        score += 2
    if meeting_context_id(meeting) == last_meeting_id:
        score += 4
    return score


def format_meeting_answer(meeting: dict, question: str = "") -> str:
    actions = meeting.get("actions", []) or []
    action_lines = []
    for action in actions:
        action_lines.append(
            f"- {normalize_value(action.get('text'))} | owner: {normalize_value(action.get('owner'), 'Not stated')} | "
            f"department: {normalize_value(action.get('department') or action.get('company'), 'Not stated')} | "
            f"deadline: {normalize_value(action.get('deadline'), 'None')} | status: {normalize_status(action)}"
        )

    lines = [
        f"Meeting: {normalize_value(meeting.get('title'), 'Untitled meeting')}",
        f"Recap: {normalize_value(meeting.get('summary') or meeting.get('recaps'), 'No summary available.')}",
        f"Objective: {normalize_value(meeting.get('objective'), 'Not provided')}",
        f"Outcome: {normalize_value(meeting.get('outcome'), 'Not provided')}",
        f"Stakeholders: {join_list(meeting.get('stakeholders', []), 'None')}",
    ]
    if action_lines:
        lines.extend(["Action items:"] + action_lines)
    else:
        lines.append("Action items: None")

    if question:
        lines.append(f"Question: {question}")
    return "\n".join(lines)


def format_recap_response(meeting: dict) -> str:
    summary = normalize_value(meeting.get("summary") or meeting.get("recaps"), "No summary available.")
    actions = meeting.get("actions", []) or []
    action_lines = [
        f"- {normalize_value(action.get('text'))} | owner: {normalize_value(action.get('owner'), 'Not stated')} | "
        f"department: {normalize_value(action.get('department') or action.get('company'), 'Not stated')} | "
        f"deadline: {normalize_value(action.get('deadline'), 'None')} | status: {normalize_status(action)}"
        for action in actions
    ]

    lines = [f"Summary: {summary}"]
    if action_lines:
        lines.extend(["Action items:"] + action_lines)
    else:
        lines.append("Action items: None")
    return "\n".join(lines)


def extract_recap_from_answer(answer: str) -> str:
    text = str(answer or "").strip()
    if not text:
        return ""
    match = re.search(
        r"(?:Summary|Recap):\s*(.*?)(?:\n(?:Objective|Outcome|Stakeholders|Action items|Question):|$)",
        text,
        flags=re.IGNORECASE | re.DOTALL,
    )
    if match:
        return normalize_value(match.group(1), "").strip()
    first_line = text.split("\n", 1)[0].strip()
    return re.sub(r"^(?:Summary|Recap):\s*", "", first_line, flags=re.IGNORECASE)


def sync_generated_summary_to_meeting(meeting: dict, answer: str) -> None:
    summary_text = extract_recap_from_answer(answer)
    if not summary_text:
        return

    updated = dict(meeting)
    updated["summary"] = summary_text
    updated["recaps"] = summary_text
    updated["objective"] = normalize_value(updated.get("objective"), "") or "Review the meeting discussion and align on the next steps."
    updated["outcome"] = normalize_value(updated.get("outcome"), "") or "Meeting recap captured from the chatbot response."
    if not updated.get("user_id"):
        updated["user_id"] = normalize_chat_user_id(meeting.get("user_id", ""))
    try:
        supabase_upsert(MEETINGS_SHEET_NAME, [updated])
    except Exception:
        pass


def answer_meeting_question(question: str, meetings: list[dict]) -> tuple[str, dict | None]:
    last_meeting_id = ""
    history = load_user_history(meetings[0].get("user_id", "") if meetings else "")
    if history:
        last_meeting_id = normalize_value(history[0].get("meeting_id"), "")

    scored = []
    for meeting in meetings:
        score = score_meeting(question, meeting, last_meeting_id=last_meeting_id)
        if score > 0:
            scored.append((score, meeting))

    relevant_meetings = [meeting for _, meeting in sorted(scored, key=lambda item: item[0], reverse=True)]
    if not relevant_meetings:
        relevant_meetings = meetings[:5]

    question_lower = question.lower().strip()
    recap_question = any(
        phrase in question_lower
        for phrase in [
            "recap",
            "recaps",
            "summary",
            "summarize",
            "summarise",
            "what's the recap",
            "whats the recap",
            "what is the recap",
        ]
    )
    about_question = any(
        phrase in question_lower
        for phrase in ["about", "summary", "recap", "objective", "topic", "agenda", "discuss"]
    )
    action_question = any(keyword in question_lower for keyword in ["action", "task", "deadline", "owner", "pending", "follow up", "follow-up"])
    requested_title = extract_meeting_title_from_question(question) if (recap_question or about_question or action_question) else ""

    if is_meeting_lookup_question(question):
        return build_lookup_response(question, meetings)

    if (recap_question or about_question) and not requested_title:
        return (
            "Please include the meeting title so I can find the saved recap, for example: "
            "`what is the recap for title: UMT internship coordinators`",
            None,
        )

    if requested_title:
        title_key = normalize_title_key(requested_title)
        titled_meetings = [
            meeting
            for meeting in relevant_meetings
            if title_key and title_key in normalize_title_key(meeting.get("title"))
        ]
        if titled_meetings:
            relevant_meetings = titled_meetings

    if relevant_meetings:
        top_meeting = relevant_meetings[0]
        if recap_question or about_question or action_question:
            answer = format_recap_response(top_meeting)
            sync_generated_summary_to_meeting(top_meeting, answer)
            return answer, top_meeting

    meeting_blocks = []
    for meeting in relevant_meetings[:5]:
        meeting_blocks.append(meeting_context_text(meeting))

    ctx = "\n\n".join(meeting_blocks) if meeting_blocks else "No meeting data available."
    system = """You are MeetIQ's AI assistant.

Answer questions using the meeting data provided.
If the user asks for a recap, objective, outcome, pending actions, or solution for action items, use the data and stay anchored to the meeting context.
Be concise, practical, and business-friendly.
"""
    user_msg = f"Meeting data:\n{ctx}\n\nQuestion: {question}"
    answer = call_ollama(system, user_msg, max_tokens=300)
    answer_text = answer.strip() or "No answer generated."
    if relevant_meetings:
        sync_generated_summary_to_meeting(relevant_meetings[0], answer_text)
    return answer_text, (relevant_meetings[0] if relevant_meetings else None)


def split_telegram_message(text: str, limit: int = 3500) -> list[str]:
    text = text or ""
    if len(text) <= limit:
        return [text]

    chunks = []
    current = ""
    for paragraph in text.split("\n"):
        candidate = f"{current}\n{paragraph}".strip() if current else paragraph
        if len(candidate) > limit and current:
            chunks.append(current)
            current = paragraph
        else:
            current = candidate
    if current:
        if len(current) <= limit:
            chunks.append(current)
        else:
            for index in range(0, len(current), limit):
                chunks.append(current[index : index + limit])
    return [chunk for chunk in chunks if chunk.strip()]


def send_telegram_message(chat_id: int, text: str, reply_to_message_id: int | None = None) -> None:
    for chunk in split_telegram_message(text):
        payload = {"chat_id": chat_id, "text": chunk}
        if reply_to_message_id is not None:
            payload["reply_to_message_id"] = reply_to_message_id
        telegram_api("sendMessage", payload=payload)


def _remember_recent_key(store: set, key) -> bool:
    if key in store:
        return False
    store.add(key)
    if len(store) > _RECENT_CACHE_LIMIT:
        store.clear()
    return True


def get_chat_session_version(chat_id: int | str) -> int:
    return _CHAT_SESSION_VERSION.get(str(chat_id), 0)


def bump_chat_session_version(chat_id: int | str) -> int:
    key = str(chat_id)
    _CHAT_SESSION_VERSION[key] = _CHAT_SESSION_VERSION.get(key, 0) + 1
    return _CHAT_SESSION_VERSION[key]


def save_meeting_and_history(
    *,
    user_id: str,
    raw_text: str,
    source_name: str,
    message_text: str,
    meeting_record: dict,
) -> tuple[dict, dict]:
    meeting_row = meeting_record
    history_thread_date = today_str()
    history_thread_title = normalize_value(meeting_row.get("title"), "Untitled meeting")
    history_thread_key = chat_thread_key(user_id, history_thread_date, history_thread_title, meeting_row.get("id", ""))
    recap_answer = format_recap_response(meeting_row)
    history_row = build_history_entry(
        user_id=user_id,
        thread_key=history_thread_key,
        thread_date=history_thread_date,
        thread_title=history_thread_title,
        question=message_text,
        answer=recap_answer,
        meeting_id=meeting_row.get("id", ""),
        meeting_title=history_thread_title,
        context="Meeting",
    )

    supabase_upsert(MEETINGS_SHEET_NAME, [meeting_row])
    supabase_insert(HISTORY_SHEET_NAME, [history_row])
    return meeting_row, history_row


def process_text_submission(user_id: str, text: str, source_name: str = "Telegram message") -> str:
    recap = summarize_meeting_text(text, source_name=source_name)
    meeting_record = build_meeting_record(text, recap, user_id=user_id, source_name=source_name)
    save_meeting_and_history(
        user_id=user_id,
        raw_text=text,
        source_name=source_name,
        message_text=f"Summarize: {source_name}",
        meeting_record=meeting_record,
    )
    return format_recap_response(meeting_record)


def process_file_submission(user_id: str, uploaded_file: TelegramUpload) -> str:
    file_name = getattr(uploaded_file, "name", "upload.bin")
    lower_name = file_name.lower()
    uploaded_file.seek(0)

    if lower_name.endswith((".wav", ".mp3", ".m4a", ".ogg", ".webm", ".mp4", ".flac")):
        raw_text = transcribe_audio_file(uploaded_file, translate_to_english=True)
    else:
        raw_text = extract_text_from_document(uploaded_file)

    if not raw_text.strip():
        raise RuntimeError("Could not extract text from the uploaded file.")
    return process_text_submission(user_id, raw_text, source_name=file_name)


def should_treat_as_question(text: str) -> bool:
    lowered = text.lower().strip()
    if not lowered:
        return False

    if lowered.startswith("/ask"):
        return True

    if len(lowered) > 260:
        question_mark_count = lowered.count("?")
        sentence_count = len(split_sentences(lowered))
        if question_mark_count <= 1 and sentence_count >= 2:
            return False

    return bool(
        re.match(r"^(what|who|when|where|why|how|which|is|are|do|does|did|can|could|should|would)\b", lowered)
        or re.match(r"^(please\s+)?(summarize|summarise|recap|summary)\b", lowered)
        or re.match(r"^(what\s+did\s+we\s+decide|what\s+is\s+the\s+meeting\s+about|what's\s+the\s+recap|whats\s+the\s+recap)\b", lowered)
        or (
            lowered.endswith("?")
            and len(lowered) < 260
        )
        or any(
            marker in lowered
            for marker in [
                "recap",
                "summary",
                "pending",
                "action item",
                "action items",
                "what did we decide",
                "what is the meeting about",
            ]
        )
    )


def handle_text_message(user_id: str, text: str) -> str:
    cleaned = text.strip()
    cleaned_lower = cleaned.lower()
    greeting_words = {"hi", "hello", "hey", "start", "help", "/help"}

    if cleaned_lower.startswith("/start") or cleaned_lower in greeting_words:
        return (
            "How to use MeetIQ Bot:\n"
            "1. Send a meeting recap, transcript, PDF, DOCX, CSV, XLSX, audio, or voice note.\n"
            "2. I will generate a short summary and action items and save it to the system.\n"
            "3. To ask about a saved recap, include the meeting title, for example:\n"
            "   what is the recap for title: UMT internship coordinators\n"
            "4. To ask if we have a meeting with a company or organization, just ask naturally:\n"
            "   do we have meeting with UMP?\n"
            "5. To ask about action items, you can also use:\n"
            "   what are the action items for title: UMT internship coordinators\n"
            "6. Use /end to stop the current session."
        )

    if cleaned_lower in {"/end", "/stop", "/cancel", "end", "stop", "cancel", "start"}:
        return NO_REPLY

    if cleaned_lower.startswith("/ask"):
        cleaned = cleaned[4:].strip()

    if cleaned.startswith("/"):
        return "Unknown command. Send a meeting note, file, or ask a recap question."

    if len(cleaned) < 20 and not should_treat_as_question(cleaned):
        return "Please send the full meeting recap text, or ask a clear question about a meeting."

    if should_treat_as_question(cleaned):
        meetings = load_user_meetings(user_id)
        answer, meeting = answer_meeting_question(cleaned, meetings)
        thread_title = normalize_value(meeting.get("title"), "General") if meeting else "General"
        thread_date = normalize_value(meeting.get("date"), today_str()) if meeting else today_str()
        thread_key = chat_thread_key(user_id, thread_date, thread_title, meeting_context_id(meeting) if meeting else "")
        history_row = build_history_entry(
            user_id=user_id,
            thread_key=thread_key,
            thread_date=thread_date,
            thread_title=thread_title,
            question=cleaned,
            answer=answer,
            meeting_id=meeting_context_id(meeting) if meeting else "",
            meeting_title=thread_title,
            context="Meeting" if meeting else "General",
        )
        supabase_insert(HISTORY_SHEET_NAME, [history_row])
        return answer

    return process_text_submission(user_id, cleaned, source_name="Telegram message")


async def process_telegram_update(update: dict) -> None:
    message = update.get("message") or update.get("edited_message")
    if not message:
        return

    chat = message.get("chat") or {}
    chat_id = chat.get("id")
    from_user = message.get("from") or {}
    user_id = str(from_user.get("id", chat_id or ""))
    message_id = message.get("message_id")
    session_version = get_chat_session_version(chat_id)

    try:
        if message.get("voice"):
            voice = message["voice"]
            file_obj = download_telegram_file(voice["file_id"], "voice.ogg")
            answer = process_file_submission(user_id, file_obj)
            if get_chat_session_version(chat_id) == session_version:
                send_telegram_message(chat_id, answer, reply_to_message_id=message_id)
        elif message.get("audio"):
            audio = message["audio"]
            file_name = audio.get("file_name") or "audio.mp3"
            file_obj = download_telegram_file(audio["file_id"], file_name)
            answer = process_file_submission(user_id, file_obj)
            if get_chat_session_version(chat_id) == session_version:
                send_telegram_message(chat_id, answer, reply_to_message_id=message_id)
        elif message.get("document"):
            doc = message["document"]
            file_name = doc.get("file_name") or "document.bin"
            file_obj = download_telegram_file(doc["file_id"], file_name)
            answer = process_file_submission(user_id, file_obj)
            if get_chat_session_version(chat_id) == session_version:
                send_telegram_message(chat_id, answer, reply_to_message_id=message_id)
        elif message.get("text"):
            text = message["text"].strip()
            if text.lower() in {"/end", "/stop", "/cancel", "end", "stop", "cancel"}:
                bump_chat_session_version(chat_id)
                return

            answer = handle_text_message(user_id, text)
            if answer != NO_REPLY and get_chat_session_version(chat_id) == session_version:
                send_telegram_message(chat_id, answer, reply_to_message_id=message_id)
        else:
            if get_chat_session_version(chat_id) == session_version:
                send_telegram_message(
                    chat_id,
                    "Send text, a voice note, or a document and I will summarise it or answer meeting questions.",
                    reply_to_message_id=message_id,
                )
    except Exception as exc:
        if get_chat_session_version(chat_id) == session_version:
            send_telegram_message(chat_id, f"Sorry, I could not process that message: {exc}", reply_to_message_id=message_id)


@app.get("/health")
def health() -> dict:
    return {"ok": True}


def _validate_webhook_secret(secret: str | None) -> None:
    expected = get_telegram_webhook_secret()
    if expected and secret and secret != expected:
        raise HTTPException(status_code=403, detail="Forbidden")


async def _handle_telegram_webhook(request: Request, secret: str | None = None):
    try:
        _validate_webhook_secret(secret)

        try:
            update = await request.json()
        except Exception:
            return JSONResponse({"ok": True, "ignored": True})

        update_id = update.get("update_id")
        if update_id is not None:
            try:
                update_id = int(update_id)
            except Exception:
                update_id = None
        if update_id is not None:
            if not _remember_recent_key(_RECENT_UPDATE_IDS, update_id):
                return JSONResponse({"ok": True, "duplicate": True})

        message = update.get("message") or update.get("edited_message")
        if not message:
            return JSONResponse({"ok": True, "ignored": True})

        asyncio.create_task(process_telegram_update(update))
        return JSONResponse({"ok": True})
    except HTTPException:
        raise
    except Exception:
        return JSONResponse({"ok": True, "ignored": True})


@app.post("/telegram/webhook")
async def telegram_webhook_plain(request: Request):
    return await _handle_telegram_webhook(request, None)


@app.post("/telegram/webhook/")
async def telegram_webhook_plain_slash(request: Request):
    return await _handle_telegram_webhook(request, None)


@app.post("/telegram/webhook/{secret}")
async def telegram_webhook(secret: str, request: Request):
    return await _handle_telegram_webhook(request, secret)


@app.post("/telegram/webhook/{secret}/")
async def telegram_webhook_slash(secret: str, request: Request):
    return await _handle_telegram_webhook(request, secret)


def set_telegram_webhook(base_url: str | None = None) -> dict:
    token = get_telegram_token()
    secret = get_telegram_webhook_secret()
    if not token or not secret:
        raise RuntimeError("TELEGRAM_BOT_TOKEN and TELEGRAM_WEBHOOK_SECRET are required.")

    target = (base_url or os.getenv("TELEGRAM_WEBHOOK_URL", "")).rstrip("/")
    if not target:
        raise RuntimeError("Provide TELEGRAM_WEBHOOK_URL or pass base_url.")

    webhook_url = f"{target}/telegram/webhook/{secret}"
    payload = {"url": webhook_url}
    return telegram_api("setWebhook", payload)


if __name__ == "__main__":
    import uvicorn

    port = int(os.getenv("PORT", "8000"))
    uvicorn.run("telegram_backend:app", host="0.0.0.0", port=port, reload=False)
