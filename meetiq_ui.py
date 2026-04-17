from datetime import date, datetime

import streamlit as st

from meetiq_constants import STATUS_CFG, STATUSES
from meetiq_utils import normalize_status, normalize_value, pill, pretty_deadline


def _entity_text(item) -> str:
    if isinstance(item, dict):
        for key in ("text", "name", "title", "value"):
            value = item.get(key)
            if value:
                return str(value).strip()
        return ""
    return str(item).strip()


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


def render_action_card(action: dict, editable: bool = False, persist_callback=None) -> None:
    status = normalize_status(action)
    cfg = STATUS_CFG.get(status, STATUS_CFG["Pending"])
    owner = normalize_value(action.get("owner"), "Not stated")
    department = normalize_value(action.get("department") or action.get("company"), "Not stated")
    suggestion = action.get("suggestion", "")
    st.markdown(
        f"""
        <div class="action-card">
            <div class="action-top">
                <div class="action-title">{normalize_value(action.get('text'), 'Untitled action')}</div>
                {pill(status, cfg['color'], cfg['bg'])}
            </div>
            <div class="action-meta">Assignee: {owner} | Department: {department} | Deadline: {pretty_deadline(normalize_value(action.get('deadline'), 'None'))}</div>
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
        deadline_value = normalize_value(action.get("deadline"), "")
        deadline_mode = st.selectbox(
            "Deadline mode",
            ["No deadline", "Set deadline"],
            index=0 if deadline_value in ("", "None") else 1,
            key=f"deadline_mode_{action['id']}",
        )
        default_deadline = date.today()
        if deadline_value not in ("", "None"):
            try:
                default_deadline = datetime.strptime(deadline_value, "%Y-%m-%d").date()
            except Exception:
                default_deadline = date.today()
        edited_deadline = st.date_input(
            "Deadline",
            value=default_deadline,
            key=f"deadline_{action['id']}",
            disabled=deadline_mode == "No deadline",
        )
        new_deadline = "None" if deadline_mode == "No deadline" else edited_deadline.isoformat()
        changed = False
        if new_status != current:
            action["status"] = new_status
            changed = True
        if new_deadline != normalize_value(action.get("deadline"), "None"):
            action["deadline"] = new_deadline
            changed = True
        if changed and persist_callback:
            persist_callback()


def render_chat_bubble(role: str, text: str) -> None:
    safe_text = str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\n", "<br>")
    st.markdown(f'<div class="chat-bubble {role}">{safe_text}</div>', unsafe_allow_html=True)


def render_summary_panel(result: dict) -> None:
    nlp = result.get("nlp_pipeline", {})
    people_count = len([text for item in nlp.get("named_entities", {}).get("persons", []) if (text := _entity_text(item))])
    action_count = len(result.get("action_items", []))
    discussion_count = len(result.get("discussion_points", []))
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
        render_kpi_card("Discussion", str(discussion_count), "Main topics", "#2563eb")
    with k3:
        render_kpi_card("People", str(people_count), "People involved", "#d97706")
