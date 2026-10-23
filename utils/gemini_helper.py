"""
Gemini Helper — Uses Google Gemini AI to generate survey questions from agenda sessions.
Uses the modern `google-genai` SDK (google.genai) with automatic model discovery.
"""

import json
import re
import os
from typing import List, Dict, Optional

# Try new SDK first (google-genai), fall back to old (google-generativeai)
try:
    from google import genai as new_genai
    NEW_SDK = True
except ImportError:
    NEW_SDK = False

try:
    import google.generativeai as old_genai
    OLD_SDK = True
except ImportError:
    OLD_SDK = False


RATING_OPTIONS = ["Excellent", "Very Good", "Good", "Fair", "Poor"]

# Model preference order for new SDK (google-genai)
NEW_SDK_MODELS = [
    "gemini-2.0-flash",
    "gemini-2.0-flash-lite",
    "gemini-2.5-flash-preview-05-20",
    "gemini-1.5-flash",
    "gemini-1.5-pro",
]

# Model preference order for old SDK (google-generativeai)
OLD_SDK_MODELS = [
    "models/gemini-2.0-flash",
    "models/gemini-2.0-flash-lite",
    "models/gemini-1.5-flash-latest",
    "models/gemini-1.5-pro-latest",
]

SYSTEM_PROMPT = """You are an expert in designing professional training feedback surveys.
Your task is to analyze a workshop/training agenda and generate ONE specific, 
well-crafted feedback question per actual session.

Rules:
1. Generate EXACTLY ONE question per session provided.
2. Keep the session number and title exactly as provided.
3. Questions must be specific to the session topic — not generic.
4. Questions are always multiple-choice with options: Excellent, Very Good, Good, Fair, Poor.
5. Question format: "How would you rate the session on [specific aspect of session topic]?"
6. Return ONLY valid JSON — no markdown, no extra text.

Return format:
[
  {
    "session_no": 1,
    "session_title": "Exact Session Title",
    "question": "How would you rate the session on [specific topic aspect]?"
  },
  ...
]"""


def _get_api_key() -> Optional[str]:
    """Get Gemini API key from Streamlit secrets or environment."""
    try:
        import streamlit as st
        key = st.secrets.get("gemini", {}).get("api_key")
        if key:
            return key
        key = st.secrets.get("GEMINI_API_KEY")
        if key:
            return key
    except Exception:
        pass
    return os.environ.get("GEMINI_API_KEY")


def _parse_json_response(text: str) -> List[Dict]:
    """Extract and parse JSON from Gemini response, handling markdown code blocks."""
    text = re.sub(r"```(?:json)?\s*", "", text)
    text = re.sub(r"```\s*", "", text)
    text = text.strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass
    match = re.search(r"\[.*\]", text, re.DOTALL)
    if match:
        try:
            return json.loads(match.group())
        except json.JSONDecodeError:
            pass
    return []


def list_available_models(api_key: str) -> List[str]:
    """
    List all available Gemini models for the given API key.
    Returns model names that support generateContent.
    """
    available = []
    if NEW_SDK:
        try:
            client = new_genai.Client(api_key=api_key)
            models = client.models.list()
            for m in models:
                name = getattr(m, "name", "") or ""
                supported = getattr(m, "supported_actions", []) or []
                # Include models that support content generation
                if "generateContent" in str(supported) or not supported:
                    # Strip "models/" prefix for cleaner display
                    short = name.replace("models/", "")
                    if "gemini" in short.lower():
                        available.append(short)
        except Exception:
            pass
    return available


def _try_generate_new_sdk(api_key: str, prompt: str) -> Optional[str]:
    """Try generating content using the new google-genai SDK."""
    if not NEW_SDK:
        return None

    client = new_genai.Client(api_key=api_key)
    errors = []

    # First, try to discover available models dynamically
    dynamic_models = []
    try:
        models = client.models.list()
        for m in models:
            name = getattr(m, "name", "") or ""
            short = name.replace("models/", "")
            if "gemini" in short.lower() and "flash" in short.lower():
                dynamic_models.append(short)
    except Exception:
        pass

    # Combine: dynamic first, then our known list
    candidates = dynamic_models + [m for m in NEW_SDK_MODELS if m not in dynamic_models]

    for model_name in candidates:
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=prompt,
            )
            return response.text
        except Exception as e:
            errors.append(f"{model_name}: {e}")
            continue

    raise RuntimeError("New SDK attempts failed:\n" + "\n".join(errors[-3:]))


def _try_generate_old_sdk(api_key: str, prompt: str) -> Optional[str]:
    """Try generating content using the old google-generativeai SDK."""
    if not OLD_SDK:
        return None

    old_genai.configure(api_key=api_key)
    errors = []

    # Try to list available models
    dynamic_models = []
    try:
        for m in old_genai.list_models():
            name = getattr(m, "name", "") or ""
            if "generateContent" in getattr(m, "supported_generation_methods", []):
                short = name.replace("models/", "")
                if "gemini" in short.lower():
                    dynamic_models.append(name)
    except Exception:
        pass

    candidates = dynamic_models + OLD_SDK_MODELS if dynamic_models else OLD_SDK_MODELS

    for model_name in candidates:
        try:
            m = old_genai.GenerativeModel(model_name)
            response = m.generate_content(prompt)
            return response.text
        except Exception as e:
            errors.append(f"{model_name}: {e}")
            continue

    raise RuntimeError("Old SDK attempts failed:\n" + "\n".join(errors[-3:]))


def generate_survey_questions(
    sessions: List[Dict],
    reference_xlsx_content: Optional[str] = None,
    api_key: Optional[str] = None,
) -> List[Dict]:
    """
    Use Gemini to generate one feedback question per session.

    Args:
        sessions: List of {"session_no": int, "title": str}
        reference_xlsx_content: Optional reference form content as text
        api_key: Gemini API key (falls back to env/secrets)

    Returns:
        List of question dicts ready for the survey editor.
    """
    if not NEW_SDK and not OLD_SDK:
        raise ImportError(
            "No Gemini SDK installed. Run: pip install google-genai"
        )

    key = api_key or _get_api_key()
    if not key:
        raise ValueError(
            "Gemini API key not found. Enter it in the sidebar or set GEMINI_API_KEY in .env"
        )

    # Build prompt
    sessions_text = "\n".join(f"{s['session_no']}. {s['title']}" for s in sessions)
    reference_section = ""
    if reference_xlsx_content:
        reference_section = f"\nReference form style:\n{reference_xlsx_content[:2000]}"

    prompt = f"""{SYSTEM_PROMPT}

Workshop Agenda Sessions:
{sessions_text}
{reference_section}

Generate exactly {len(sessions)} questions, one per session, as JSON array."""

    # Try new SDK first, then old SDK
    raw_text = None
    errors = []

    if NEW_SDK:
        try:
            raw_text = _try_generate_new_sdk(key, prompt)
        except Exception as e:
            errors.append(f"[google-genai SDK] {e}")

    if raw_text is None and OLD_SDK:
        try:
            raw_text = _try_generate_old_sdk(key, prompt)
        except Exception as e:
            errors.append(f"[google-generativeai SDK] {e}")

    if raw_text is None:
        combined = " | ".join(errors)
        raise RuntimeError(f"Gemini API error: {combined}")

    parsed = _parse_json_response(raw_text)
    if not parsed:
        return _generate_fallback_questions(sessions)

    result = []
    for item in parsed:
        if isinstance(item, dict):
            result.append({
                "session_no": item.get("session_no", 0),
                "session_title": item.get("session_title", ""),
                "question": item.get("question", ""),
            })

    return result if result else _generate_fallback_questions(sessions)


def _generate_fallback_questions(sessions: List[Dict]) -> List[Dict]:
    """Generate basic fallback questions without AI."""
    return [
        {
            "session_no": s["session_no"],
            "session_title": s["title"],
            "question": f"How would you rate the session on {s['title']}?",
        }
        for s in sessions
    ]


def build_full_survey(ai_questions: List[Dict]) -> List[Dict]:
    """
    Build the complete survey structure from AI-generated questions.
    Prepends participant fields and appends overall + suggestions questions.
    """
    survey = []
    q_num = 1

    # --- Section 1: Participant Information ---
    participant_fields = [
        {"field": "Full Name", "qtype": "Short Answer", "required": True},
        {"field": "Designation", "qtype": "Short Answer", "required": True},
        {"field": "Organisation/Department", "qtype": "Short Answer", "required": True},
    ]
    for pf in participant_fields:
        survey.append({
            "no": q_num,
            "section": "Participant Information",
            "session_no": "",
            "session": "",
            "category": "Participant Info",
            "question": pf["field"],
            "question_type": pf["qtype"],
            "options": "",
            "required": "Yes",
        })
        q_num += 1

    # --- Section 2: Session Feedback ---
    options_str = " | ".join(RATING_OPTIONS)
    for q in ai_questions:
        survey.append({
            "no": q_num,
            "section": "Session Feedback",
            "session_no": str(q.get("session_no", "")),
            "session": q.get("session_title", ""),
            "category": "Session Rating",
            "question": q.get("question", ""),
            "question_type": "Multiple Choice",
            "options": options_str,
            "required": "Yes",
        })
        q_num += 1

    # --- Section 3: Overall Workshop ---
    survey.append({
        "no": q_num,
        "section": "Overall Workshop",
        "session_no": "",
        "session": "",
        "category": "Overall Rating",
        "question": "How would you rate the overall workshop?",
        "question_type": "Multiple Choice",
        "options": options_str,
        "required": "Yes",
    })
    q_num += 1

    # --- Section 4: Suggestions ---
    survey.append({
        "no": q_num,
        "section": "Suggestions",
        "session_no": "",
        "session": "",
        "category": "Open Feedback",
        "question": "Please provide your suggestions for improving future workshops.",
        "question_type": "Paragraph",
        "options": "",
        "required": "No",
    })

    return survey
