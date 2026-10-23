"""
AI Survey Generator — Main Streamlit Application

Transforms a workshop agenda (DOCX) into a complete, editable feedback survey
using Gemini AI, with Excel export and Google Forms creation via OAuth 2.0.
"""

import os
import json
import time
import copy
from io import BytesIO
from typing import List, Dict, Optional

import streamlit as st
from dotenv import load_dotenv

# Load local .env file (for development)
load_dotenv()

# ─── Page Configuration ────────────────────────────────────────────────────────
st.set_page_config(
    page_title="AI Survey Generator",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={
        "About": "AI-powered workshop feedback survey generator using Google Gemini.",
    },
)

# ─── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* ── Google Fonts ── */
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap');

/* ── Root variables ── */
:root {
    --primary: #6366f1;
    --primary-dark: #4f46e5;
    --primary-light: #818cf8;
    --secondary: #06b6d4;
    --accent: #f59e0b;
    --success: #10b981;
    --danger: #ef4444;
    --warning: #f59e0b;
    --bg-dark: #0f0f1a;
    --bg-card: #1a1a2e;
    --bg-card2: #16213e;
    --border: rgba(99, 102, 241, 0.2);
    --text-muted: #94a3b8;
    --radius: 16px;
    --radius-sm: 8px;
    --shadow: 0 8px 32px rgba(0,0,0,0.3);
    --glow: 0 0 20px rgba(99,102,241,0.3);
}

/* ── Global ── */
html, body, [class*="css"] {
    font-family: 'Inter', sans-serif !important;
}

.main .block-container {
    padding: 2rem 2.5rem !important;
    max-width: 1400px;
}

/* ── Dark background ── */
.stApp {
    background: linear-gradient(135deg, #0f0f1a 0%, #1a1a2e 50%, #16213e 100%) !important;
    min-height: 100vh;
}

/* ── Hero header ── */
.hero-header {
    background: linear-gradient(135deg, rgba(99,102,241,0.15) 0%, rgba(6,182,212,0.1) 100%);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 2.5rem 3rem;
    margin-bottom: 2rem;
    position: relative;
    overflow: hidden;
    backdrop-filter: blur(10px);
}
.hero-header::before {
    content: '';
    position: absolute;
    top: -50%;
    right: -20%;
    width: 500px;
    height: 500px;
    background: radial-gradient(circle, rgba(99,102,241,0.12) 0%, transparent 70%);
    pointer-events: none;
}
.hero-title {
    font-size: 2.6rem;
    font-weight: 800;
    background: linear-gradient(135deg, #6366f1, #06b6d4, #f59e0b);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
    margin: 0 0 0.5rem 0;
    line-height: 1.2;
}
.hero-subtitle {
    font-size: 1.05rem;
    color: var(--text-muted);
    margin: 0;
    font-weight: 400;
}
.hero-badges {
    display: flex;
    gap: 0.5rem;
    margin-top: 1rem;
    flex-wrap: wrap;
}
.badge {
    background: rgba(99,102,241,0.15);
    border: 1px solid rgba(99,102,241,0.3);
    color: #818cf8;
    padding: 0.25rem 0.75rem;
    border-radius: 999px;
    font-size: 0.75rem;
    font-weight: 600;
    letter-spacing: 0.05em;
}
.badge.cyan { background: rgba(6,182,212,0.12); border-color: rgba(6,182,212,0.3); color: #22d3ee; }
.badge.amber { background: rgba(245,158,11,0.12); border-color: rgba(245,158,11,0.3); color: #fbbf24; }
.badge.green { background: rgba(16,185,129,0.12); border-color: rgba(16,185,129,0.3); color: #34d399; }

/* ── Cards ── */
.glass-card {
    background: rgba(26,26,46,0.8);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 1.75rem;
    margin-bottom: 1.25rem;
    backdrop-filter: blur(12px);
    box-shadow: var(--shadow);
    transition: all 0.3s ease;
}
.glass-card:hover {
    border-color: rgba(99,102,241,0.4);
    box-shadow: var(--glow), var(--shadow);
}
.card-title {
    font-size: 1.1rem;
    font-weight: 700;
    color: #e2e8f0;
    margin: 0 0 1rem 0;
    display: flex;
    align-items: center;
    gap: 0.5rem;
}
.card-title .icon {
    font-size: 1.2rem;
}

/* ── Step indicator ── */
.step-container {
    display: flex;
    align-items: center;
    gap: 0;
    margin-bottom: 2rem;
    overflow-x: auto;
    padding: 0.5rem 0;
}
.step-item {
    display: flex;
    align-items: center;
    flex: 1;
    min-width: 0;
}
.step-circle {
    width: 40px;
    height: 40px;
    border-radius: 50%;
    display: flex;
    align-items: center;
    justify-content: center;
    font-weight: 700;
    font-size: 0.9rem;
    flex-shrink: 0;
    transition: all 0.3s ease;
    border: 2px solid;
}
.step-circle.done {
    background: linear-gradient(135deg, #10b981, #059669);
    border-color: #10b981;
    color: white;
}
.step-circle.active {
    background: linear-gradient(135deg, #6366f1, #4f46e5);
    border-color: #6366f1;
    color: white;
    box-shadow: 0 0 16px rgba(99,102,241,0.5);
}
.step-circle.inactive {
    background: rgba(30,30,60,0.8);
    border-color: rgba(99,102,241,0.25);
    color: #64748b;
}
.step-label {
    font-size: 0.72rem;
    font-weight: 600;
    text-align: center;
    margin-top: 0.4rem;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
    max-width: 80px;
}
.step-label.active { color: #818cf8; }
.step-label.done { color: #34d399; }
.step-label.inactive { color: #475569; }
.step-connector {
    height: 2px;
    flex: 1;
    background: linear-gradient(90deg, rgba(99,102,241,0.4), rgba(99,102,241,0.1));
    margin: 0 0.5rem;
    margin-bottom: 1.4rem;
}
.step-connector.done {
    background: linear-gradient(90deg, #10b981, #34d399);
}
.step-wrapper {
    display: flex;
    flex-direction: column;
    align-items: center;
}

/* ── Info/Alert boxes ── */
.info-box {
    background: rgba(99,102,241,0.08);
    border-left: 3px solid #6366f1;
    border-radius: 0 var(--radius-sm) var(--radius-sm) 0;
    padding: 0.85rem 1.2rem;
    margin-bottom: 1rem;
    color: #c7d2fe;
    font-size: 0.9rem;
    line-height: 1.6;
}
.success-box {
    background: rgba(16,185,129,0.08);
    border-left: 3px solid #10b981;
    border-radius: 0 var(--radius-sm) var(--radius-sm) 0;
    padding: 0.85rem 1.2rem;
    margin-bottom: 1rem;
    color: #6ee7b7;
    font-size: 0.9rem;
}
.warning-box {
    background: rgba(245,158,11,0.08);
    border-left: 3px solid #f59e0b;
    border-radius: 0 var(--radius-sm) var(--radius-sm) 0;
    padding: 0.85rem 1.2rem;
    margin-bottom: 1rem;
    color: #fcd34d;
    font-size: 0.9rem;
}
.error-box {
    background: rgba(239,68,68,0.08);
    border-left: 3px solid #ef4444;
    border-radius: 0 var(--radius-sm) var(--radius-sm) 0;
    padding: 0.85rem 1.2rem;
    margin-bottom: 1rem;
    color: #fca5a5;
    font-size: 0.9rem;
}

/* ── Stat pills ── */
.stat-row {
    display: flex;
    gap: 1rem;
    flex-wrap: wrap;
    margin-bottom: 1.5rem;
}
.stat-pill {
    background: rgba(99,102,241,0.1);
    border: 1px solid rgba(99,102,241,0.2);
    border-radius: 999px;
    padding: 0.4rem 1.1rem;
    display: flex;
    align-items: center;
    gap: 0.4rem;
    font-size: 0.82rem;
    color: #c7d2fe;
    font-weight: 500;
}
.stat-pill .num {
    font-weight: 700;
    color: #a5b4fc;
    font-size: 0.9rem;
}

/* ── Section headers in table ── */
.section-tag {
    display: inline-block;
    padding: 0.15rem 0.6rem;
    border-radius: 999px;
    font-size: 0.72rem;
    font-weight: 600;
}
.section-tag.participant { background: rgba(99,102,241,0.2); color: #a5b4fc; }
.section-tag.session { background: rgba(16,185,129,0.2); color: #6ee7b7; }
.section-tag.overall { background: rgba(245,158,11,0.2); color: #fcd34d; }
.section-tag.suggestions { background: rgba(239,68,68,0.2); color: #fca5a5; }

/* ── Sidebar ── */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0f0f1a 0%, #1a1a2e 100%) !important;
    border-right: 1px solid var(--border);
}
[data-testid="stSidebar"] .block-container {
    padding: 1.5rem 1rem !important;
}

/* ── Buttons ── */
.stButton > button {
    background: linear-gradient(135deg, #6366f1 0%, #4f46e5 100%) !important;
    color: white !important;
    border: none !important;
    border-radius: var(--radius-sm) !important;
    font-family: 'Inter', sans-serif !important;
    font-weight: 600 !important;
    font-size: 0.9rem !important;
    padding: 0.6rem 1.4rem !important;
    transition: all 0.25s ease !important;
    box-shadow: 0 4px 12px rgba(99,102,241,0.3) !important;
    letter-spacing: 0.02em !important;
}
.stButton > button:hover {
    background: linear-gradient(135deg, #4f46e5 0%, #4338ca 100%) !important;
    box-shadow: 0 6px 20px rgba(99,102,241,0.45) !important;
    transform: translateY(-1px) !important;
}
.stButton > button:active {
    transform: translateY(0) !important;
}

/* ── Download buttons ── */
.stDownloadButton > button {
    background: linear-gradient(135deg, #10b981 0%, #059669 100%) !important;
    color: white !important;
    border: none !important;
    border-radius: var(--radius-sm) !important;
    font-weight: 600 !important;
    padding: 0.6rem 1.4rem !important;
    transition: all 0.25s ease !important;
    box-shadow: 0 4px 12px rgba(16,185,129,0.3) !important;
}
.stDownloadButton > button:hover {
    transform: translateY(-1px) !important;
    box-shadow: 0 6px 20px rgba(16,185,129,0.45) !important;
}

/* ── File uploader ── */
[data-testid="stFileUploader"] {
    border: 2px dashed rgba(99,102,241,0.3) !important;
    border-radius: var(--radius) !important;
    background: rgba(99,102,241,0.04) !important;
    transition: all 0.2s ease;
}
[data-testid="stFileUploader"]:hover {
    border-color: rgba(99,102,241,0.6) !important;
    background: rgba(99,102,241,0.08) !important;
}

/* ── Data editor ── */
.stDataFrame, [data-testid="stDataFrame"] {
    border-radius: var(--radius) !important;
    overflow: hidden !important;
    border: 1px solid var(--border) !important;
}

/* ── Text inputs ── */
.stTextInput > div > div > input,
.stTextArea > div > div > textarea,
.stSelectbox > div > div > div {
    background: rgba(30,30,60,0.6) !important;
    border: 1px solid rgba(99,102,241,0.2) !important;
    border-radius: var(--radius-sm) !important;
    color: #e2e8f0 !important;
    font-family: 'Inter', sans-serif !important;
}
.stTextInput > div > div > input:focus,
.stTextArea > div > div > textarea:focus {
    border-color: #6366f1 !important;
    box-shadow: 0 0 0 2px rgba(99,102,241,0.2) !important;
}

/* ── Expander ── */
.streamlit-expanderHeader {
    background: rgba(26,26,46,0.6) !important;
    border-radius: var(--radius-sm) !important;
    color: #c7d2fe !important;
    font-weight: 600 !important;
}

/* ── Tabs ── */
.stTabs [data-baseweb="tab-list"] {
    background: rgba(26,26,46,0.6) !important;
    border-radius: var(--radius-sm) !important;
    gap: 0.25rem;
    padding: 0.25rem;
}
.stTabs [data-baseweb="tab"] {
    border-radius: 6px !important;
    color: #94a3b8 !important;
    font-weight: 500 !important;
}
.stTabs [aria-selected="true"] {
    background: linear-gradient(135deg, #6366f1, #4f46e5) !important;
    color: white !important;
}

/* ── Spinner ── */
.stSpinner > div {
    border-color: #6366f1 !important;
}

/* ── Dividers ── */
hr {
    border-color: var(--border) !important;
    margin: 1.5rem 0 !important;
}

/* ── Form submit button ── */
.form-url-box {
    background: rgba(16,185,129,0.08);
    border: 1px solid rgba(16,185,129,0.3);
    border-radius: var(--radius);
    padding: 1.5rem;
    margin-top: 1rem;
}
.form-url-box h4 {
    color: #34d399;
    margin: 0 0 0.75rem;
    font-size: 1rem;
    font-weight: 700;
}
.form-url-box a {
    color: #6ee7b7 !important;
    word-break: break-all;
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.85rem;
}

/* ── Floating action area ── */
.action-bar {
    position: sticky;
    bottom: 0;
    background: linear-gradient(0deg, rgba(15,15,26,0.95) 0%, transparent 100%);
    padding: 1rem 0 0.5rem;
    z-index: 100;
}

/* ── Scrollbar ── */
::-webkit-scrollbar { width: 6px; height: 6px; }
::-webkit-scrollbar-track { background: rgba(30,30,60,0.3); }
::-webkit-scrollbar-thumb { background: rgba(99,102,241,0.4); border-radius: 3px; }
::-webkit-scrollbar-thumb:hover { background: rgba(99,102,241,0.7); }

/* ── Metric cards ── */
[data-testid="stMetric"] {
    background: rgba(26,26,46,0.6) !important;
    border: 1px solid var(--border) !important;
    border-radius: var(--radius-sm) !important;
    padding: 0.75rem 1rem !important;
}
[data-testid="stMetricValue"] {
    color: #818cf8 !important;
    font-weight: 700 !important;
}
[data-testid="stMetricLabel"] {
    color: var(--text-muted) !important;
}

/* ── Sidebar logo/brand ── */
.sidebar-brand {
    display: flex;
    align-items: center;
    gap: 0.75rem;
    padding: 0.75rem 0 1.5rem;
    border-bottom: 1px solid var(--border);
    margin-bottom: 1.5rem;
}
.sidebar-brand-icon {
    width: 42px;
    height: 42px;
    background: linear-gradient(135deg, #6366f1, #06b6d4);
    border-radius: 10px;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 1.3rem;
}
.sidebar-brand-name {
    font-size: 1rem;
    font-weight: 800;
    color: #e2e8f0;
    line-height: 1.1;
}
.sidebar-brand-tagline {
    font-size: 0.72rem;
    color: #6366f1;
    font-weight: 500;
}

/* ── Progress bar custom ── */
.progress-bar-outer {
    background: rgba(99,102,241,0.1);
    border-radius: 999px;
    height: 6px;
    margin: 0.5rem 0 1rem;
    overflow: hidden;
}
.progress-bar-inner {
    height: 100%;
    background: linear-gradient(90deg, #6366f1, #06b6d4);
    border-radius: 999px;
    transition: width 0.5s ease;
}

/* ── Question row cards ── */
.q-row {
    display: flex;
    align-items: flex-start;
    gap: 1rem;
    padding: 1rem;
    background: rgba(22,22,50,0.7);
    border: 1px solid rgba(99,102,241,0.12);
    border-radius: var(--radius-sm);
    margin-bottom: 0.6rem;
    transition: all 0.2s ease;
}
.q-row:hover {
    border-color: rgba(99,102,241,0.3);
    background: rgba(30,30,60,0.8);
}
.q-num {
    width: 28px;
    height: 28px;
    border-radius: 50%;
    background: linear-gradient(135deg, rgba(99,102,241,0.3), rgba(99,102,241,0.1));
    border: 1px solid rgba(99,102,241,0.4);
    color: #a5b4fc;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 0.75rem;
    font-weight: 700;
    flex-shrink: 0;
}
.q-content { flex: 1; min-width: 0; }
.q-title {
    color: #e2e8f0;
    font-size: 0.9rem;
    font-weight: 500;
    margin-bottom: 0.25rem;
    line-height: 1.4;
}
.q-meta {
    color: #64748b;
    font-size: 0.75rem;
    display: flex;
    gap: 0.75rem;
    flex-wrap: wrap;
}
</style>
""", unsafe_allow_html=True)


# ─── Session State Initialization ──────────────────────────────────────────────
def init_session_state():
    defaults = {
        "step": 1,
        "agenda_sessions": [],
        "survey_rows": [],
        "form_title": "Workshop Feedback Survey",
        "reference_content": None,
        "oauth_token": None,
        "form_edit_url": None,
        "form_published_url": None,
        "gemini_api_key": "",
        "generation_done": False,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_session_state()


# ─── Helper Functions ──────────────────────────────────────────────────────────
def render_step_indicator(current_step: int):
    steps = [
        ("📤", "Upload"),
        ("🤖", "Generate"),
        ("✏️", "Review"),
        ("📤", "Export"),
    ]

    html_parts = ['<div class="step-container">']
    for i, (icon, label) in enumerate(steps, start=1):
        if i < current_step:
            circle_cls, label_cls = "done", "done"
            circle_content = "✓"
        elif i == current_step:
            circle_cls, label_cls = "active", "active"
            circle_content = str(i)
        else:
            circle_cls, label_cls = "inactive", "inactive"
            circle_content = str(i)

        html_parts.append(f'''
        <div class="step-item">
            <div class="step-wrapper">
                <div class="step-circle {circle_cls}">{circle_content}</div>
                <div class="step-label {label_cls}">{label}</div>
            </div>
        </div>''')

        if i < len(steps):
            conn_cls = "done" if i < current_step else ""
            html_parts.append(f'<div class="step-connector {conn_cls}"></div>')

    html_parts.append('</div>')
    st.markdown("".join(html_parts), unsafe_allow_html=True)


def render_hero():
    st.markdown("""
    <div class="hero-header">
        <h1 class="hero-title">📋 AI Survey Generator</h1>
        <p class="hero-subtitle">Transform workshop agendas into professional feedback surveys instantly</p>
        <div class="hero-badges">
            <span class="badge">✨ Gemini AI</span>
            <span class="badge cyan">🔗 Google Forms</span>
            <span class="badge amber">📊 Excel Export</span>
            <span class="badge green">🔒 OAuth 2.0</span>
        </div>
    </div>
    """, unsafe_allow_html=True)


def get_gemini_key() -> str:
    """Get Gemini API key from session state, secrets, or env."""
    if st.session_state.get("gemini_api_key"):
        return st.session_state["gemini_api_key"]
    try:
        key = st.secrets.get("gemini", {}).get("api_key", "")
        if key and key != "YOUR_GEMINI_API_KEY_HERE":
            return key
    except Exception:
        pass
    return os.environ.get("GEMINI_API_KEY", "")


def count_survey_stats(rows: List[Dict]) -> Dict:
    """Count questions by section."""
    stats = {"total": 0, "sessions": 0, "participant": 0, "overall": 0}
    for row in rows:
        stats["total"] += 1
        s = row.get("section", "")
        if s == "Session Feedback":
            stats["sessions"] += 1
        elif s == "Participant Information":
            stats["participant"] += 1
        elif s in ("Overall Workshop", "Suggestions"):
            stats["overall"] += 1
    return stats


# ─── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div class="sidebar-brand">
        <div class="sidebar-brand-icon">📋</div>
        <div>
            <div class="sidebar-brand-name">AI Survey<br>Generator</div>
            <div class="sidebar-brand-tagline">Powered by Gemini</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("### ⚙️ Configuration")

    # Gemini API Key
    with st.expander("🔑 Gemini API Key", expanded=not bool(get_gemini_key())):
        api_input = st.text_input(
            "Gemini API Key",
            value=st.session_state.get("gemini_api_key", ""),
            type="password",
            placeholder="AIza...",
            help="Get your key at https://aistudio.google.com/app/apikey",
            key="api_key_input",
            label_visibility="collapsed",
        )
        if api_input:
            st.session_state["gemini_api_key"] = api_input
        if get_gemini_key():
            st.markdown('<div class="success-box">✅ API Key configured</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="warning-box">⚠️ Gemini API key required for AI generation</div>', unsafe_allow_html=True)
        st.markdown("[Get API Key →](https://aistudio.google.com/app/apikey)", unsafe_allow_html=False)

    st.divider()

    # Form title
    st.markdown("### 📝 Survey Settings")
    form_title_input = st.text_input(
        "Survey Title",
        value=st.session_state.get("form_title", "Workshop Feedback Survey"),
        key="form_title_input",
    )
    st.session_state["form_title"] = form_title_input

    st.divider()

    # Progress
    step = st.session_state["step"]
    progress_pct = int((step - 1) / 3 * 100)
    st.markdown("### 📊 Progress")
    st.markdown(f"""
    <div class="progress-bar-outer">
        <div class="progress-bar-inner" style="width:{progress_pct}%;"></div>
    </div>
    <div style="color:#64748b;font-size:0.8rem;text-align:right;">Step {step} of 4</div>
    """, unsafe_allow_html=True)

    if st.session_state["survey_rows"]:
        stats = count_survey_stats(st.session_state["survey_rows"])
        st.markdown(f"""
        <div class="stat-row" style="flex-direction:column;">
            <div class="stat-pill">📋 Total Questions: <span class="num">{stats['total']}</span></div>
            <div class="stat-pill">🎯 Session Questions: <span class="num">{stats['sessions']}</span></div>
        </div>
        """, unsafe_allow_html=True)

    st.divider()

    # Navigation buttons
    st.markdown("### 🧭 Navigation")
    col_back, col_fwd = st.columns(2)
    with col_back:
        if st.button("← Back", disabled=step == 1, use_container_width=True, key="nav_back"):
            st.session_state["step"] = max(1, step - 1)
            st.rerun()
    with col_fwd:
        can_advance = (
            (step == 1 and bool(st.session_state["agenda_sessions"])) or
            (step == 2 and bool(st.session_state["survey_rows"])) or
            (step == 3 and bool(st.session_state["survey_rows"])) or
            step == 4
        )
        if st.button("Next →", disabled=not can_advance or step == 4, use_container_width=True, key="nav_fwd"):
            st.session_state["step"] = min(4, step + 1)
            st.rerun()

    st.divider()
    st.markdown("""
    <div style="color:#475569;font-size:0.72rem;text-align:center;line-height:1.6;">
        AI Survey Generator<br>
        Built with Streamlit + Gemini<br>
        <a href="https://github.com/mohit-dhoundiyal/AI-Survey-Generator" 
           style="color:#6366f1;">GitHub Repository ↗</a>
    </div>
    """, unsafe_allow_html=True)


# ─── Main Content ──────────────────────────────────────────────────────────────
render_hero()
render_step_indicator(st.session_state["step"])

# ══════════════════════════════════════════════════════════════
# STEP 1: UPLOAD
# ══════════════════════════════════════════════════════════════
if st.session_state["step"] == 1:
    st.markdown('<div class="glass-card"><p class="card-title"><span class="icon">📤</span> Step 1 — Upload Your Workshop Agenda</p>', unsafe_allow_html=True)

    st.markdown("""
    <div class="info-box">
        Upload a DOCX file containing your workshop or training agenda. 
        The AI will identify the actual sessions and generate one targeted feedback question per session.
        Optionally upload an existing feedback form (XLSX) as a style reference.
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns([3, 2], gap="large")

    with col1:
        st.markdown("#### 📄 Workshop Agenda (Required)")
        docx_file = st.file_uploader(
            "Upload DOCX agenda",
            type=["docx"],
            key="docx_upload",
            help="Your workshop/training agenda in Microsoft Word format",
            label_visibility="collapsed",
        )

        if docx_file:
            st.markdown(f"""
            <div class="success-box">
                ✅ <strong>{docx_file.name}</strong> uploaded successfully 
                ({docx_file.size / 1024:.1f} KB)
            </div>
            """, unsafe_allow_html=True)

            # Parse the DOCX
            try:
                from utils.agenda_parser import parse_agenda_docx
                sessions = parse_agenda_docx(docx_file.read())
                docx_file.seek(0)

                if sessions:
                    st.session_state["agenda_sessions"] = sessions
                    st.session_state["generation_done"] = False

                    st.markdown(f"""
                    <div class="success-box">
                        🎯 Found <strong>{len(sessions)}</strong> session(s) in the agenda
                    </div>
                    """, unsafe_allow_html=True)

                    with st.expander(f"👁️ Preview extracted sessions ({len(sessions)} found)", expanded=True):
                        for s in sessions:
                            st.markdown(f"""
                            <div class="q-row">
                                <div class="q-num">{s['session_no']}</div>
                                <div class="q-content">
                                    <div class="q-title">{s['title']}</div>
                                </div>
                            </div>
                            """, unsafe_allow_html=True)
                else:
                    st.markdown("""
                    <div class="warning-box">
                        ⚠️ No sessions detected automatically. The document may use an unusual format.
                        Try the manual text entry option below.
                    </div>
                    """, unsafe_allow_html=True)
            except ImportError:
                st.markdown("""
                <div class="error-box">
                    ❌ python-docx not installed. Run: <code>pip install python-docx</code>
                </div>
                """, unsafe_allow_html=True)
            except Exception as e:
                st.markdown(f"""
                <div class="error-box">❌ Error parsing DOCX: {str(e)}</div>
                """, unsafe_allow_html=True)

    with col2:
        st.markdown("#### 📊 Reference Form (Optional)")
        xlsx_file = st.file_uploader(
            "Upload XLSX reference",
            type=["xlsx"],
            key="xlsx_upload",
            help="An existing feedback form to use as style reference",
            label_visibility="collapsed",
        )
        if xlsx_file:
            st.markdown(f"""
            <div class="success-box">
                ✅ Reference form: <strong>{xlsx_file.name}</strong>
            </div>
            """, unsafe_allow_html=True)
            try:
                from utils.excel_exporter import read_xlsx_as_text
                ref_content = read_xlsx_as_text(xlsx_file.read())
                st.session_state["reference_content"] = ref_content
                xlsx_file.seek(0)
            except Exception:
                st.session_state["reference_content"] = None

        st.markdown("---")
        st.markdown("#### ✏️ Or Enter Agenda Manually")
        manual_text = st.text_area(
            "Paste agenda items (one per line)",
            height=180,
            placeholder="Session 1: Introduction to AI\nSession 2: Machine Learning Basics\nSession 3: Deep Learning\n...",
            key="manual_agenda_input",
            label_visibility="collapsed",
        )
        if st.button("📋 Parse Manual Text", key="parse_manual", use_container_width=True):
            if manual_text.strip():
                try:
                    from utils.agenda_parser import parse_agenda_text
                    sessions = parse_agenda_text(manual_text)
                    if sessions:
                        st.session_state["agenda_sessions"] = sessions
                        st.session_state["generation_done"] = False
                        st.success(f"✅ Extracted {len(sessions)} sessions from text")
                        st.rerun()
                    else:
                        st.warning("No sessions detected in the text provided.")
                except Exception as e:
                    st.error(f"Error: {e}")
            else:
                st.warning("Please enter some agenda text first.")

    st.markdown('</div>', unsafe_allow_html=True)

    # Advance button
    if st.session_state["agenda_sessions"]:
        st.markdown('<div style="margin-top:1rem;">', unsafe_allow_html=True)
        if st.button("🤖 Proceed to AI Generation →", key="go_step2", type="primary", use_container_width=False):
            st.session_state["step"] = 2
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════
# STEP 2: AI GENERATION
# ══════════════════════════════════════════════════════════════
elif st.session_state["step"] == 2:
    st.markdown('<div class="glass-card"><p class="card-title"><span class="icon">🤖</span> Step 2 — AI Survey Generation</p>', unsafe_allow_html=True)

    sessions = st.session_state["agenda_sessions"]

    st.markdown(f"""
    <div class="stat-row">
        <div class="stat-pill">📋 Agenda Sessions: <span class="num">{len(sessions)}</span></div>
        <div class="stat-pill">✨ AI Model: <span class="num">Gemini 2.0 Flash (auto)</span></div>
    </div>
    """, unsafe_allow_html=True)

    # Show sessions to process
    with st.expander("📋 Sessions to process", expanded=False):
        for s in sessions:
            st.markdown(f"""
            <div class="q-row">
                <div class="q-num">{s['session_no']}</div>
                <div class="q-content">
                    <div class="q-title">{s['title']}</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

    if st.session_state.get("generation_done") and st.session_state.get("survey_rows"):
        st.markdown("""
        <div class="success-box">
            ✅ Survey questions have already been generated! Proceed to the Review step 
            or click Regenerate to create new questions.
        </div>
        """, unsafe_allow_html=True)

        stats = count_survey_stats(st.session_state["survey_rows"])
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Questions", stats["total"])
        col2.metric("Session Questions", stats["sessions"])
        col3.metric("Fixed Questions", stats["participant"] + stats["overall"])

    # Generation controls
    api_key = get_gemini_key()
    has_key = bool(api_key)

    # ── Test API key / list models ──
    if has_key:
        with st.expander("🔍 Test API Key & Discover Available Models", expanded=False):
            if st.button("🧪 Test API Key", key="test_api_key"):
                with st.spinner("Querying Gemini API for available models..."):
                    try:
                        from utils.gemini_helper import list_available_models
                        found = list_available_models(api_key)
                        if found:
                            st.markdown(f"""
                            <div class="success-box">
                                ✅ API key works! Found <strong>{len(found)}</strong> available model(s):
                            </div>
                            """, unsafe_allow_html=True)
                            for m in found:
                                st.markdown(f"- `{m}`")
                        else:
                            st.markdown("""
                            <div class="warning-box">
                                ⚠️ API key connected but no models listed. Will try defaults on generation.
                            </div>
                            """, unsafe_allow_html=True)
                    except Exception as e:
                        st.markdown(f"""
                        <div class="error-box">
                            ❌ API key test failed: {str(e)}<br><br>
                            <strong>Common causes:</strong><br>
                            • Invalid or expired API key<br>
                            • Key not enabled for Gemini API<br>
                            • Get a fresh key at <a href="https://aistudio.google.com/app/apikey" target="_blank">aistudio.google.com</a>
                        </div>
                        """, unsafe_allow_html=True)

    col_gen, col_skip = st.columns([2, 1], gap="medium")

    with col_gen:
        if not has_key:
            st.markdown("""
            <div class="warning-box">
                ⚠️ Gemini API key not configured. You can still generate a basic survey without AI 
                by clicking "Generate Without AI" below.
            </div>
            """, unsafe_allow_html=True)

        btn_label = "✨ Generate Survey with Gemini AI" if has_key else "✨ Generate Survey (No AI — Basic Mode)"
        if st.button(btn_label, key="generate_btn", type="primary", use_container_width=True):
            with st.spinner("🤖 Gemini is analyzing your agenda and crafting questions..."):
                try:
                    if has_key:
                        from utils.gemini_helper import generate_survey_questions, build_full_survey
                        ai_qs = generate_survey_questions(
                            sessions,
                            reference_xlsx_content=st.session_state.get("reference_content"),
                            api_key=api_key,
                        )
                    else:
                        from utils.gemini_helper import _generate_fallback_questions, build_full_survey
                        ai_qs = _generate_fallback_questions(sessions)

                    survey = build_full_survey(ai_qs)
                    st.session_state["survey_rows"] = survey
                    st.session_state["generation_done"] = True
                    time.sleep(0.5)
                    st.success(f"✅ Generated {len(survey)} questions successfully!")
                    st.rerun()
                except Exception as e:
                    err_msg = str(e)
                    st.markdown(f"""
                    <div class="error-box">
                        ❌ <strong>Generation failed</strong><br><br>
                        <code>{err_msg}</code><br><br>
                        <strong>Try these fixes:</strong><br>
                        1. Use the <em>"Test API Key"</em> expander above to verify your key<br>
                        2. Make sure your key is from <a href="https://aistudio.google.com/app/apikey" target="_blank">Google AI Studio</a> (not Google Cloud)<br>
                        3. Or click <em>"Generate Without AI"</em> below for basic mode
                    </div>
                    """, unsafe_allow_html=True)

        # Basic mode fallback button always visible
        if has_key:
            if st.button("⚡ Generate Without AI (Basic Mode)", key="generate_basic", use_container_width=True):
                with st.spinner("Generating basic survey..."):
                    from utils.gemini_helper import _generate_fallback_questions, build_full_survey
                    ai_qs = _generate_fallback_questions(sessions)
                    survey = build_full_survey(ai_qs)
                    st.session_state["survey_rows"] = survey
                    st.session_state["generation_done"] = True
                    st.success(f"✅ Basic survey generated ({len(survey)} questions). You can edit questions in Step 3.")
                    st.rerun()

    with col_skip:
        if st.session_state.get("survey_rows"):
            if st.button("✏️ Go to Review →", key="go_review", use_container_width=True):
                st.session_state["step"] = 3
                st.rerun()


# ══════════════════════════════════════════════════════════════
# STEP 3: REVIEW & EDIT
# ══════════════════════════════════════════════════════════════
elif st.session_state["step"] == 3:
    st.markdown('<div class="glass-card"><p class="card-title"><span class="icon">✏️</span> Step 3 — Review & Edit Survey</p>', unsafe_allow_html=True)

    survey_rows = st.session_state.get("survey_rows", [])
    stats = count_survey_stats(survey_rows)

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Questions", stats["total"])
    col2.metric("Session Feedback", stats["sessions"])
    col3.metric("Participant Info", stats["participant"])
    col4.metric("Overall + Tips", stats["overall"])

    st.markdown("""
    <div class="info-box">
        💡 You can edit questions below by clicking on any cell. 
        Add custom questions using the form at the bottom, or delete individual questions.
        Changes are saved automatically.
    </div>
    """, unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

    # ── Edit table ──
    tab_editor, tab_add, tab_preview = st.tabs(["📝 Edit Questions", "➕ Add Question", "👁️ Preview"])

    with tab_editor:
        if not survey_rows:
            st.warning("No survey questions yet. Go back to Step 2 to generate them.")
        else:
            import pandas as pd

            df = pd.DataFrame(survey_rows)

            # Allow editing all fields except "no"
            editable_cols = ["question", "section", "session_no", "session", "question_type", "options", "required"]
            col_config = {
                "no": st.column_config.NumberColumn("No.", width="small"),
                "section": st.column_config.SelectboxColumn(
                    "Section",
                    options=["Participant Information", "Session Feedback", "Overall Workshop", "Suggestions"],
                    width="medium",
                ),
                "session_no": st.column_config.TextColumn("Session No.", width="small"),
                "session": st.column_config.TextColumn("Session Title", width="medium"),
                "category": st.column_config.TextColumn("Category", width="small"),
                "question": st.column_config.TextColumn("Question", width="large"),
                "question_type": st.column_config.SelectboxColumn(
                    "Type",
                    options=["Short Answer", "Paragraph", "Multiple Choice"],
                    width="small",
                ),
                "options": st.column_config.TextColumn("Options (| separated)", width="medium"),
                "required": st.column_config.SelectboxColumn(
                    "Required",
                    options=["Yes", "No"],
                    width="small",
                ),
            }

            edited_df = st.data_editor(
                df,
                use_container_width=True,
                num_rows="dynamic",
                column_config=col_config,
                key="survey_editor",
                hide_index=True,
            )

            col_save, col_reset = st.columns([1, 1])
            with col_save:
                if st.button("💾 Save Changes", key="save_edits", type="primary"):
                    # Renumber rows
                    edited_records = edited_df.to_dict("records")
                    for i, row in enumerate(edited_records, start=1):
                        row["no"] = i
                    st.session_state["survey_rows"] = edited_records
                    st.success("✅ Changes saved!")
                    st.rerun()
            with col_reset:
                if st.button("🔄 Regenerate Questions", key="regen"):
                    st.session_state["generation_done"] = False
                    st.session_state["step"] = 2
                    st.rerun()

    with tab_add:
        st.markdown("#### Add a Custom Question")
        with st.form("add_question_form", clear_on_submit=True):
            new_section = st.selectbox(
                "Section",
                ["Participant Information", "Session Feedback", "Overall Workshop", "Suggestions"],
                key="new_section",
            )
            new_session_no = st.text_input("Session No. (optional)", placeholder="e.g., 5", key="new_session_no")
            new_session_title = st.text_input("Session Title (optional)", placeholder="e.g., Panel Discussion", key="new_session_title")
            new_question = st.text_area("Question *", placeholder="How would you rate...", key="new_question", height=80)
            new_qtype = st.selectbox(
                "Question Type",
                ["Multiple Choice", "Short Answer", "Paragraph"],
                key="new_qtype",
            )
            new_options = st.text_input(
                "Options (| separated, for Multiple Choice)",
                value="Excellent | Very Good | Good | Fair | Poor",
                key="new_options",
            )
            new_required = st.radio("Required?", ["Yes", "No"], horizontal=True, key="new_required")

            if st.form_submit_button("➕ Add Question", type="primary"):
                if not new_question.strip():
                    st.error("Question text is required.")
                else:
                    current_rows = st.session_state["survey_rows"]
                    new_no = len(current_rows) + 1
                    current_rows.append({
                        "no": new_no,
                        "section": new_section,
                        "session_no": new_session_no,
                        "session": new_session_title,
                        "category": "Custom",
                        "question": new_question.strip(),
                        "question_type": new_qtype,
                        "options": new_options.strip(),
                        "required": new_required,
                    })
                    st.session_state["survey_rows"] = current_rows
                    st.success(f"✅ Question added! Total: {len(current_rows)}")
                    st.rerun()

    with tab_preview:
        st.markdown("#### 📋 Survey Preview")
        st.markdown("*This is how the survey will appear to participants.*")

        current_section = None
        for row in st.session_state.get("survey_rows", []):
            sec = row.get("section", "")
            if sec != current_section:
                current_section = sec
                section_cls_map = {
                    "Participant Information": "participant",
                    "Session Feedback": "session",
                    "Overall Workshop": "overall",
                    "Suggestions": "suggestions",
                }
                sec_cls = section_cls_map.get(sec, "session")
                st.markdown(f"""
                <div style="margin:1.5rem 0 0.5rem;">
                    <span class="section-tag {sec_cls}">{sec}</span>
                </div>
                """, unsafe_allow_html=True)

            req_badge = "🔴" if row.get("required") == "Yes" else "⚪"
            qtype = row.get("question_type", "")
            session_label = ""
            if row.get("session"):
                session_label = f' <span style="color:#64748b;font-size:0.8rem;">[{row["session"]}]</span>'

            st.markdown(f"""
            <div class="q-row">
                <div class="q-num">{row.get('no', '')}</div>
                <div class="q-content">
                    <div class="q-title">{req_badge} {row.get('question', '')}{session_label}</div>
                    <div class="q-meta">
                        <span>📌 {qtype}</span>
                        {'<span>📎 ' + row.get('options','') + '</span>' if row.get('options') else ''}
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)

    # Bottom navigation
    st.markdown('<div style="margin-top:1.5rem;">', unsafe_allow_html=True)
    if st.button("📤 Proceed to Export →", key="go_export", type="primary"):
        st.session_state["step"] = 4
        st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════
# STEP 4: EXPORT
# ══════════════════════════════════════════════════════════════
elif st.session_state["step"] == 4:
    st.markdown('<div class="glass-card"><p class="card-title"><span class="icon">📤</span> Step 4 — Export Your Survey</p>', unsafe_allow_html=True)

    survey_rows = st.session_state.get("survey_rows", [])
    form_title = st.session_state.get("form_title", "Workshop Feedback Survey")

    if not survey_rows:
        st.markdown("""
        <div class="error-box">
            ❌ No survey questions found. Please go back to Step 2 to generate questions first.
        </div>
        """, unsafe_allow_html=True)
        if st.button("← Go Back to Step 2"):
            st.session_state["step"] = 2
            st.rerun()
    else:
        stats = count_survey_stats(survey_rows)
        st.markdown(f"""
        <div class="stat-row">
            <div class="stat-pill">📋 Total: <span class="num">{stats['total']} questions</span></div>
            <div class="stat-pill">🎯 Sessions: <span class="num">{stats['sessions']}</span></div>
            <div class="stat-pill">📄 Form Title: <span class="num">{form_title}</span></div>
        </div>
        """, unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

        # Export tabs
        tab_excel, tab_script, tab_gform = st.tabs([
            "📊 Excel Export",
            "📜 Apps Script",
            "🔗 Google Forms (OAuth)",
        ])

        # ── EXCEL TAB ──
        with tab_excel:
            st.markdown('<div class="glass-card">', unsafe_allow_html=True)
            st.markdown("#### 📊 Download Excel Workbook")
            st.markdown("""
            <div class="info-box">
                Downloads a styled Excel file with a single <strong>Feedback Form</strong> worksheet.
                Columns include: No., Section, Session No., Session, Category, Question, 
                Question Type, Options, Required.
            </div>
            """, unsafe_allow_html=True)

            try:
                from utils.excel_exporter import export_to_excel
                excel_bytes = export_to_excel(survey_rows)
                safe_title = form_title.replace(" ", "_").replace("/", "-")

                st.download_button(
                    label="⬇️ Download Excel (.xlsx)",
                    data=excel_bytes,
                    file_name=f"{safe_title}_Feedback_Survey.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_excel",
                    use_container_width=True,
                    type="primary",
                )
                st.markdown("""
                <div class="success-box">
                    ✅ Excel file ready for download. The workbook includes a styled header row, 
                    color-coded sections, and frozen panes for easy navigation.
                </div>
                """, unsafe_allow_html=True)
            except Exception as e:
                st.markdown(f'<div class="error-box">❌ Excel export error: {str(e)}</div>', unsafe_allow_html=True)

            # Show column structure
            with st.expander("📋 Excel Column Structure"):
                col_data = {
                    "Column": ["No.", "Section", "Session No.", "Session", "Category",
                               "Question", "Question Type", "Options", "Required"],
                    "Description": [
                        "Sequential question number",
                        "Survey section (e.g., Session Feedback)",
                        "Workshop session number",
                        "Workshop session title",
                        "Question category",
                        "The actual feedback question text",
                        "Short Answer / Multiple Choice / Paragraph",
                        "Answer choices separated by |",
                        "Yes or No",
                    ]
                }
                import pandas as pd
                st.dataframe(pd.DataFrame(col_data), use_container_width=True, hide_index=True)

            st.markdown('</div>', unsafe_allow_html=True)

        # ── APPS SCRIPT TAB ──
        with tab_script:
            st.markdown('<div class="glass-card">', unsafe_allow_html=True)
            st.markdown("#### 📜 Google Apps Script Fallback")
            st.markdown("""
            <div class="info-box">
                Download a ready-to-use Google Apps Script that creates your Google Form 
                directly from your Google account — no OAuth setup needed in the app.
            </div>
            """, unsafe_allow_html=True)

            try:
                from utils.apps_script_generator import generate_apps_script
                script_content = generate_apps_script(survey_rows, form_title=form_title)

                st.download_button(
                    label="⬇️ Download Apps Script (.gs)",
                    data=script_content.encode("utf-8"),
                    file_name=f"{form_title.replace(' ', '_')}_SurveyScript.gs",
                    mime="text/plain",
                    key="dl_script",
                    use_container_width=True,
                )

                with st.expander("📖 How to use this script"):
                    st.markdown("""
                    1. **Download** the `.gs` file above
                    2. Go to [script.google.com](https://script.google.com)
                    3. Create a **New Project**
                    4. **Paste** the entire script content
                    5. Click **Run** → `createSurveyForm`
                    6. **Authorize** with your Google account
                    7. The Form URL will appear in the **Execution Log**
                    """)

                with st.expander("👁️ Script Preview"):
                    st.code(script_content[:2000] + ("\n..." if len(script_content) > 2000 else ""), language="javascript")

            except Exception as e:
                st.markdown(f'<div class="error-box">❌ Script generation error: {str(e)}</div>', unsafe_allow_html=True)

            st.markdown('</div>', unsafe_allow_html=True)

        # ── GOOGLE FORMS TAB ──
        with tab_gform:
            st.markdown('<div class="glass-card">', unsafe_allow_html=True)
            st.markdown("#### 🔗 Create Google Form Directly via OAuth 2.0")

            # Check for OAuth config
            try:
                from utils.google_forms import get_oauth_config
                oauth_cfg = get_oauth_config()
                has_oauth = bool(oauth_cfg.get("client_id") and oauth_cfg["client_id"] != "YOUR_GOOGLE_CLIENT_ID.apps.googleusercontent.com")
            except Exception:
                has_oauth = False
                oauth_cfg = {}

            if not has_oauth:
                st.markdown("""
                <div class="warning-box">
                    ⚠️ Google OAuth not configured. To use this feature, you need to:
                    <ol style="margin:0.5rem 0 0 1rem;">
                        <li>Create a Google Cloud project</li>
                        <li>Enable the Google Forms API</li>
                        <li>Create OAuth 2.0 credentials (Web Application)</li>
                        <li>Add credentials to <code>.streamlit/secrets.toml</code></li>
                    </ol>
                    <br>Use the <strong>Apps Script tab</strong> as a working alternative.
                </div>
                """, unsafe_allow_html=True)

                with st.expander("📖 Setup Instructions for Google OAuth"):
                    st.markdown("""
                    **1. Google Cloud Console Setup**
                    ```
                    → console.cloud.google.com
                    → Create/select a project
                    → APIs & Services → Library → Enable "Google Forms API"
                    → APIs & Services → OAuth consent screen → External
                    → APIs & Services → Credentials → Create → OAuth Client ID → Web Application
                    ```
                    
                    **2. Add Redirect URI**
                    ```
                    http://localhost:8501/component/streamlit_oauth.authorize_button
                    ```
                    
                    **3. Update `.streamlit/secrets.toml`**
                    ```toml
                    [google_oauth]
                    client_id = "YOUR_CLIENT_ID.apps.googleusercontent.com"
                    client_secret = "YOUR_CLIENT_SECRET"
                    redirect_uri = "http://localhost:8501/component/streamlit_oauth.authorize_button"
                    ```
                    
                    **4. Restart the app** and come back to this tab.
                    """)
            else:
                st.markdown("""
                <div class="info-box">
                    ✅ Google OAuth is configured! Click below to authorize and create your Google Form directly.
                    The form will appear in the authorizing Google account's Drive.
                </div>
                """, unsafe_allow_html=True)

                # Check if already have a token
                if st.session_state.get("oauth_token"):
                    st.markdown("""
                    <div class="success-box">✅ Google account authorized.</div>
                    """, unsafe_allow_html=True)

                    if not (st.session_state.get("form_edit_url") and st.session_state.get("form_published_url")):
                        if st.button("🚀 Create Google Form Now", key="create_form", type="primary", use_container_width=True):
                            with st.spinner("Creating your Google Form..."):
                                try:
                                    from utils.google_forms import create_google_form
                                    edit_url, published_url = create_google_form(
                                        access_token=st.session_state["oauth_token"],
                                        survey_rows=survey_rows,
                                        form_title=form_title,
                                    )
                                    st.session_state["form_edit_url"] = edit_url
                                    st.session_state["form_published_url"] = published_url
                                    st.rerun()
                                except Exception as e:
                                    st.markdown(f'<div class="error-box">❌ Form creation failed: {str(e)}</div>', unsafe_allow_html=True)

                    if st.session_state.get("form_edit_url"):
                        st.markdown(f"""
                        <div class="form-url-box">
                            <h4>🎉 Google Form Created Successfully!</h4>
                            <p style="color:#94a3b8;font-size:0.85rem;margin-bottom:0.5rem;">Edit URL:</p>
                            <a href="{st.session_state['form_edit_url']}" target="_blank">
                                {st.session_state['form_edit_url']}
                            </a>
                            <p style="color:#94a3b8;font-size:0.85rem;margin:0.75rem 0 0.5rem;">Published URL (share with participants):</p>
                            <a href="{st.session_state['form_published_url']}" target="_blank">
                                {st.session_state['form_published_url']}
                            </a>
                        </div>
                        """, unsafe_allow_html=True)

                        col_edit, col_share = st.columns(2)
                        with col_edit:
                            st.link_button("✏️ Open Form Editor", st.session_state["form_edit_url"], use_container_width=True)
                        with col_share:
                            st.link_button("🔗 Open Published Form", st.session_state["form_published_url"], use_container_width=True)
                else:
                    # Show OAuth button
                    try:
                        from streamlit_oauth import authorize_button

                        st.markdown("**Click below to sign in with your Google account:**")
                        result = authorize_button(
                            name="Sign in with Google",
                            icon="https://upload.wikimedia.org/wikipedia/commons/c/c1/Google_%22G%22_logo.svg",
                            redirect_uri=oauth_cfg.get("redirect_uri", ""),
                            client_id=oauth_cfg.get("client_id", ""),
                            client_secret=oauth_cfg.get("client_secret", ""),
                            scope="https://www.googleapis.com/auth/forms.body openid email",
                            key="google_oauth_btn",
                        )

                        if result and result.get("token"):
                            token_data = result["token"]
                            access_token = (
                                token_data.get("access_token")
                                if isinstance(token_data, dict)
                                else str(token_data)
                            )
                            if access_token:
                                st.session_state["oauth_token"] = access_token
                                st.rerun()

                    except ImportError:
                        st.markdown("""
                        <div class="error-box">
                            ❌ streamlit-oauth not installed. Run:<br>
                            <code>pip install streamlit-oauth</code>
                        </div>
                        """, unsafe_allow_html=True)
                    except Exception as e:
                        st.markdown(f'<div class="error-box">❌ OAuth error: {str(e)}</div>', unsafe_allow_html=True)

                if st.session_state.get("oauth_token"):
                    if st.button("🔓 Disconnect Google Account", key="logout_google"):
                        st.session_state["oauth_token"] = None
                        st.session_state["form_edit_url"] = None
                        st.session_state["form_published_url"] = None
                        st.rerun()

            st.markdown('</div>', unsafe_allow_html=True)

        # ── Start Over button ──
        st.markdown("---")
        col_a, col_b = st.columns([1, 3])
        with col_a:
            if st.button("🔄 Start Over", key="start_over"):
                for key in list(st.session_state.keys()):
                    del st.session_state[key]
                st.rerun()
        with col_b:
            if st.button("← Back to Review", key="back_to_review"):
                st.session_state["step"] = 3
                st.rerun()


# ─── Footer ────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown("""
<div style="text-align:center;color:#475569;font-size:0.8rem;padding:1rem 0;">
    🤖 AI Survey Generator &nbsp;|&nbsp; Built with 
    <a href="https://streamlit.io" style="color:#6366f1;">Streamlit</a> + 
    <a href="https://ai.google.dev/" style="color:#6366f1;">Gemini AI</a> &nbsp;|&nbsp;
    <a href="https://github.com/mohit-dhoundiyal/AI-Survey-Generator" style="color:#6366f1;">View on GitHub ↗</a>
</div>
""", unsafe_allow_html=True)
