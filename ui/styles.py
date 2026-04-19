import streamlit as st


APP_STYLES = """
<style>
:root {
    --bg-outer: #0E1F2F;
    --surface: #ffffff;
    --border: #d8dceb;
    --text: #0f172a;
    --text-soft: #6e7f96;
    --brand: #0E1B48;
    --brand-2: #27425D;
    --soft-blush: #E2CAD8;
}
.stApp {
    background:
        radial-gradient(circle at top right, rgba(135, 167, 208, 0.14), transparent 30%),
        radial-gradient(circle at bottom left, rgba(193, 141, 180, 0.18), transparent 36%),
        var(--bg-outer);
}
.block-container {
    padding: 1.25rem 1.5rem 2rem;
    margin: 1rem auto;
    max-width: 1380px;
    background: linear-gradient(180deg, var(--surface) 0%, #fcfbfe 100%);
    border-radius: 28px;
    box-shadow: 0 22px 60px rgba(15, 23, 42, 0.18);
}
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #27425D 0%, #0E1B48 52%, #0E1F2F 100%);
}
.sidebar-title {
    margin: 0 0 1rem;
    color: #ffffff !important;
    font-size: 1.45rem;
    font-weight: 800;
}
.sidebar-subtitle {
    margin: -0.7rem 0 1rem;
    color: rgba(255,255,255,0.78);
    font-size: 0.88rem;
}
.hero-panel, .kpi-card, .action-card, .info-card, .section-card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 18px;
    box-shadow: 0 14px 28px rgba(15, 23, 42, 0.06);
}
.hero-panel { padding: 1.35rem 1.4rem; }
.hero-badge {
    display: inline-block;
    background: var(--soft-blush);
    color: var(--brand);
    padding: 0.35rem 0.7rem;
    border-radius: 999px;
    font-size: 0.8rem;
    font-weight: 700;
    margin-bottom: 0.8rem;
}
.hero-grid { display:grid; grid-template-columns: repeat(2,minmax(0,1fr)); gap:.9rem; margin-top:1.1rem; }
.kpi-card, .completion-card, .info-card, .section-card { padding: 1rem 1.05rem; }
.kpi-label { color: var(--text-soft); font-size: 0.86rem; text-transform: uppercase; letter-spacing: 0.08em; font-weight: 700; }
.kpi-value { margin-top: 0.45rem; font-size: 1.85rem; line-height: 1; font-weight: 800; color: var(--text); }
.kpi-subtitle, .mini-copy { color: var(--text-soft); }
.action-card { padding: 1rem 1rem 0.85rem; margin-bottom: 0.75rem; }
.action-top { display:flex; align-items:start; justify-content:space-between; gap:.8rem; margin-bottom:.45rem; }
.action-title { color: var(--text); font-weight: 700; font-size: 1rem; }
.action-meta { color: #27425D; font-size: 0.92rem; margin-bottom: 0.28rem; }
.chat-thread { display:flex; flex-direction:column; gap:.85rem; margin:.9rem 0 1rem; }
.chat-bubble { max-width:min(82%,820px); padding:.9rem 1rem; border-radius:18px; border:1px solid var(--border); line-height:1.55; font-size:.98rem; white-space:pre-wrap; }
.chat-bubble.user { margin-left:auto; background:linear-gradient(135deg,#27425D,#0E1B48); color:#fff; border-bottom-right-radius:6px; }
.chat-bubble.assistant { margin-right:auto; background:#f7f4f8; color:var(--text); border-bottom-left-radius:6px; }
.summary-section { background:#fff; border:1px solid var(--border); border-radius:16px; padding:1rem; min-height:140px; }
.summary-section-title { font-size:.9rem; font-weight:700; letter-spacing:.08em; text-transform:uppercase; color:#27425D; margin-bottom:.55rem; }
.summary-section-body { color:#0f172a; line-height:1.6; }
</style>
"""


def inject_styles():
    st.markdown(APP_STYLES, unsafe_allow_html=True)
