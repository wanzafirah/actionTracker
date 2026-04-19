import ast
import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.services import call_ollama
from meetiq_utils import better_objective_from_transcript, compact_transcript_for_prompt, fallback_action_items, fallback_discussion_points, fallback_key_decisions, join_list, load_text_list, looks_like_copied_intro, normalize_status, normalize_value, parse_yes_no, summary_needs_expansion


PIPELINE_SYSTEM = """You are a meeting intelligence system. Return ONLY valid JSON.

Goals:
- Extract persons, organizations, dates, and locations.
- Identify the meeting objective from the transcript and metadata.
- Write a detailed but readable 6-8 sentence summary that stays faithful to the transcript.
- Preserve the original meeting context, key participants, decisions, concerns, commitments, and next steps.
- Avoid generic filler such as "hello everyone" unless it is truly relevant to the meeting.
- Extract key decisions, discussion points, and action items.
- Mark follow-up as true only when something is still pending.
- The objective should be a concise paraphrase of the meeting purpose, not a copy of the opening recap line.
- The summary should synthesize the whole discussion in your own words; do not restate the first transcript sentence verbatim.
- For long discussions, include the main context, key points, decisions, requests, and pending items instead of repeating the meeting intro.

Rules:
- Treat only explicitly stated tasks, requests, assignments, and pending items as action items.
- Do not infer hidden or implied tasks from general discussion or meeting purpose.
- If a recap says a partner or external party must provide missing details before the initiative can proceed, capture the follow-up as a TalentCorp action item to request and coordinate that information.
- Keep the action item specific to the concrete next step, not a generic restatement of the meeting topic.
- Only keep action items that belong to TalentCorp or TalentCorp internal teams. If the assignee or responsibility is clearly for another organization, do not place it in action_items unless the action is specifically a TalentCorp follow-up to collect or coordinate that external information.
- External-party responsibilities can still appear in the summary or discussion points, but not as TalentCorp action items.
- Always return at least 1-3 discussion_points when the transcript contains actual meeting content.
- discussion_points should capture the main topics discussed, reviewed, aligned, explored, or presented.
- Return key_decisions only when the transcript clearly states a decision, agreement, confirmation, or approval.
- If a recap includes important dates, months, launch periods, deadlines, or timelines, include them only when they are explicitly mentioned.
- If owner or deadline is missing, use "Not stated" and "None".
- Prefer separate action items instead of merging unrelated tasks, but only when each task is explicitly stated.
- If structured metadata is provided, use it as context.
- Use metadata such as stakeholders, organizations, departments, and report-by names to resolve who the meeting is about and what follow-up is needed.
- If the transcript is a long discussion, summarize the full conversation into the main topics, what was agreed, what was requested, and what remains pending.
- Keep the summary concrete and specific. Mention the actual organizations, speakers, projects, timelines, and requests when they are present.
- For each action item, include a practical suggestion that helps the owner solve or move the task forward.

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


def transcript_sentences(text: str) -> list:
    raw_sentences = re.split(r"(?<=[.!?])\s+|\n+", text)
    return [sentence.strip(" -•\t") for sentence in raw_sentences if sentence and sentence.strip()]


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
    repair_system = "You repair malformed meeting-analysis JSON. Return only valid JSON with the same meaning."
    repair_prompt = f"Fix this into valid JSON only:\n\n{raw[:3000]}"
    repaired = call_ollama(repair_system, repair_prompt, max_tokens=300)
    return extract_json(repaired)


def build_safe_pipeline_result(transcript: str, metadata: dict | None = None) -> dict:
    metadata = metadata or {}
    discussion_points = fallback_discussion_points(transcript)
    key_decisions = fallback_key_decisions(transcript)
    title = normalize_value(metadata.get("Title"), "") or normalize_value(metadata.get("Activity Title"), "") or "Untitled"
    meeting_type = normalize_value(metadata.get("Activity Type"), "") or "Not Provided"
    category = normalize_value(metadata.get("Category"), "") or "Not Provided"
    objective = discussion_points[0] if discussion_points else "Objective not clearly extracted."
    return {
        "title": title,
        "meeting_type": meeting_type,
        "category": category,
        "nlp_pipeline": {
            "token_count": 0,
            "sentence_count": len(transcript_sentences(transcript)),
            "named_entities": {"persons": [], "organizations": [], "dates": [], "locations": []},
        },
        "classification": {
            "action_items_count": 0,
            "decisions_count": len(key_decisions),
            "discussion_points_count": len(discussion_points),
        },
        "objective": objective,
        "summary": "",
        "outcome": "Not provided",
        "follow_up": False,
        "follow_up_reason": "",
        "key_decisions": key_decisions,
        "discussion_points": discussion_points,
        "action_items": [],
        "estimated_budget": 0,
        "budget_notes": "",
    }


def normalize_result(result: dict, transcript: str, metadata: dict | None = None) -> dict:
    safe = build_safe_pipeline_result(transcript, metadata)
    if not isinstance(result, dict):
        raise ValueError("Pipeline result is not a dictionary.")
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


def refine_summary_fields(result: dict, transcript: str, metadata: dict | None = None) -> dict:
    metadata = metadata or {}
    summary_text = normalize_value(result.get("summary"), "")
    objective_text = normalize_value(result.get("objective"), "")
    should_refine_summary = summary_needs_expansion(summary_text, transcript)
    should_refine_objective = not objective_text or looks_like_copied_intro(objective_text, transcript)
    if not should_refine_summary and not should_refine_objective:
        return result

    metadata_lines = [f"{label}: {normalize_value(value, '')}" for label, value in metadata.items() if normalize_value(value, "")]
    action_lines = []
    for action in result.get("action_items", []):
        if isinstance(action, dict):
            action_lines.append(
                " | ".join(
                    [
                        normalize_value(action.get("text"), ""),
                        f"owner={normalize_value(action.get('owner'), 'Not stated')}",
                        f"department={normalize_value(action.get('department') or action.get('company'), 'Not stated')}",
                        f"deadline={normalize_value(action.get('deadline'), 'None')}",
                    ]
                )
            )

    refine_system = """You rewrite meeting recaps into strong business-friendly prose.
Return ONLY valid JSON with keys: summary, objective.
Rules:
- summary must be a 6-8 sentence synthesis of the full meeting.
- objective must be 1-2 sentences describing the actual purpose of the meeting.
- Do not copy the opening line.
- Read the whole context and explain what happened, what was agreed, what was requested, and what remains pending.
- Keep names, organizations, projects, dates, and next steps accurate.
"""
    refine_prompt = (
        f"Metadata:\n{chr(10).join(metadata_lines) or 'None'}\n\n"
        f"Discussion points:\n{join_list(result.get('discussion_points', []), 'None')}\n\n"
        f"Key decisions:\n{join_list(result.get('key_decisions', []), 'None')}\n\n"
        f"Action items:\n{chr(10).join(action_lines) or 'None'}\n\n"
        f"Transcript:\n{transcript[:5000]}"
    )
    refined = extract_json(call_ollama(refine_system, refine_prompt, max_tokens=900))

    if should_refine_summary:
        refined_summary = normalize_value(refined.get("summary"), "")
        if refined_summary and not looks_like_copied_intro(refined_summary, transcript):
            result["summary"] = refined_summary
    if should_refine_objective:
        refined_objective = normalize_value(refined.get("objective"), "")
        if refined_objective and not looks_like_copied_intro(refined_objective, transcript):
            result["objective"] = refined_objective
        elif not objective_text:
            result["objective"] = better_objective_from_transcript(transcript) or "Objective not clearly extracted."
    return result


def _meeting_context_id(meeting: dict) -> str:
    return normalize_value(meeting.get("meetingID") or meeting.get("activityId") or meeting.get("id"), "")


def _meeting_context_text(meeting: dict) -> str:
    action_texts = []
    for action in meeting.get("actions", []):
        if isinstance(action, dict):
            action_texts.append(
                " ".join(
                    [
                        normalize_value(action.get("text"), ""),
                        normalize_value(action.get("owner"), ""),
                        normalize_value(action.get("department") or action.get("company"), ""),
                        normalize_value(action.get("deadline"), ""),
                        normalize_value(action.get("status"), ""),
                        normalize_value(action.get("suggestion"), ""),
                    ]
                )
            )
    return " ".join(
        [
            normalize_value(meeting.get("title"), ""),
            normalize_value(meeting.get("summary"), ""),
            normalize_value(meeting.get("recaps"), ""),
            normalize_value(meeting.get("objective"), ""),
            normalize_value(meeting.get("outcome"), ""),
            normalize_value(meeting.get("meetingID"), ""),
            normalize_value(meeting.get("activityId"), ""),
            normalize_value(meeting.get("deptName"), ""),
            normalize_value(meeting.get("department"), ""),
            normalize_value(meeting.get("sltdepartment"), ""),
            join_list(meeting.get("stakeholders", []), ""),
            join_list(meeting.get("companies", []), ""),
            join_list(meeting.get("discussionPoints", []), ""),
            join_list(meeting.get("keyDecisions", []), ""),
            " ".join(action_texts),
        ]
    ).lower()


def _format_meeting_answer(meeting: dict, question_lower: str) -> str:
    title = normalize_value(meeting.get("title"), "this meeting")
    summary = normalize_value(meeting.get("summary") or meeting.get("recaps"), "")
    objective = normalize_value(meeting.get("objective"), "")
    outcome = normalize_value(meeting.get("outcome"), "")
    discussions = load_text_list(meeting.get("discussionPoints", []))
    decisions = load_text_list(meeting.get("keyDecisions", []))
    actions = meeting.get("actions", [])
    wants_solution = any(marker in question_lower for marker in ["solution", "solve", "how to", "what should", "what can", "next step", "next steps", "fix", "resolve"])

    parts = [f'Meeting: "{title}"']
    if summary:
        parts.append(f"Recap: {summary}")
    elif objective:
        parts.append(f"Objective: {objective}")
    elif discussions:
        parts.append("Discussion points:")
        parts.extend(f"- {item}" for item in discussions[:5])
    if outcome and outcome.lower() not in {"none", "not provided"}:
        parts.append(f"Outcome: {outcome}")
    if decisions:
        parts.append("Key decisions:")
        parts.extend(f"- {item}" for item in decisions[:5])
    if actions:
        parts.append("Action items:")
        for action in actions:
            line = f"- {normalize_value(action.get('text'), 'Untitled action')} | owner: {normalize_value(action.get('owner'), 'Not stated')} | status: {normalize_status(action)} | deadline: {normalize_value(action.get('deadline'), 'None')}"
            suggestion = normalize_value(action.get("suggestion"), "")
            if wants_solution and suggestion:
                line += f" | solution: {suggestion}"
            parts.append(line)
    return "\n".join(parts)


def run_pipeline(transcript: str, metadata: dict | None = None) -> dict:
    cleaned_transcript = compact_transcript_for_prompt(transcript.strip(), max_chars=2200)
    objective_only = False
    metadata_lines = [f"{label}: {normalize_value(value, '')}" for label, value in (metadata or {}).items() if normalize_value(value, "")]
    metadata_block = "\n".join(metadata_lines)
    objective_note = "This transcript is objective-only. Return no action items and set follow_up to false.\n" if objective_only else ""
    user_msg = (
        "Return detailed JSON with summary, objective, outcome, follow-up, action items, deadlines, and solution suggestions.\n"
        "Use only tasks and deadlines that are explicitly stated in the meeting recap.\n"
        "Make the summary detailed enough that someone who missed the meeting can understand the context, discussion, agreements, pending items, and next steps.\n"
        "For each action item, include owner, department or company, deadline, and a practical suggestion.\n"
        f"{objective_note}"
        f"Activity metadata:\n{metadata_block or 'None provided'}\n\n"
        f"Meeting content:\n{cleaned_transcript}"
    )
    raw = call_ollama(PIPELINE_SYSTEM, user_msg, max_tokens=1800)
    try:
        result = extract_json(raw)
    except json.JSONDecodeError:
        result = recover_json_with_ollama(raw)
    except Exception:
        result = recover_json_with_ollama(raw)

    result = normalize_result(result, cleaned_transcript, metadata)
    if not result.get("discussion_points"):
        result["discussion_points"] = fallback_discussion_points(cleaned_transcript)
    if not result.get("key_decisions"):
        result["key_decisions"] = fallback_key_decisions(cleaned_transcript)
    if not result.get("action_items") and not objective_only:
        result["action_items"] = fallback_action_items(cleaned_transcript)
    if result.get("action_items"):
        result["follow_up"] = True
        if not normalize_value(result.get("follow_up_reason"), ""):
            result["follow_up_reason"] = "Action items are still pending."
    result.setdefault("classification", {})
    result["classification"]["action_items_count"] = len(result.get("action_items", []))
    result["classification"]["decisions_count"] = len(result.get("key_decisions", []))
    result["classification"]["discussion_points_count"] = len(result.get("discussion_points", []))
    return refine_summary_fields(result, cleaned_transcript, metadata)


def chat_with_meetings(question: str, meetings: list) -> str:
    question_lower = question.lower().strip()
    question_tokens = {
        token
        for token in re.findall(r"[a-zA-Z0-9&]+", question_lower)
        if len(token) >= 3 and token not in {"what", "when", "where", "which", "there", "their", "about", "with", "from", "have", "that", "this", "item", "items", "task", "tasks", "action", "actions", "meeting", "meetings"}
    }

    scored_meetings = []
    for meeting in meetings:
        blob = _meeting_context_text(meeting)
        score = sum(1 for token in question_tokens if token in blob)
        if score > 0:
            scored_meetings.append((score, meeting))

    relevant_meetings = [meeting for _, meeting in sorted(scored_meetings, key=lambda item: item[0], reverse=True)] or meetings[:5]
    action_question = any(keyword in question_lower for keyword in ["action", "task", "deadline", "owner", "pending", "follow up", "follow-up"])
    about_question = any(phrase in question_lower for phrase in ["about", "summary", "objective", "what is the meeting", "what was the meeting", "topic", "agenda", "discuss"])

    if (action_question or about_question) and relevant_meetings:
        return _format_meeting_answer(relevant_meetings[0], question_lower)

    ctx = "\n\n".join(_format_meeting_answer(meeting, question_lower) for meeting in relevant_meetings[:5]) if relevant_meetings else "No meeting data available."
    system = """You are MeetIQ's AI assistant.
Answer questions using stored meeting data whenever the question is about meetings, tasks, follow-up items, decisions, owners, history, recap details, or action-item solutions.
When answering from meeting data, mention the relevant meeting title, owner, deadline, and status when available.
Be concise, practical, and business-friendly.
"""
    user_msg = f"Meeting data:\n{ctx}\n\nQuestion: {question}"
    return call_ollama(system, user_msg, max_tokens=400)
