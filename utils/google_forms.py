"""
Google Forms Integration — Creates a Google Form using the Forms API via OAuth 2.0.
"""

import os
from typing import List, Dict, Optional, Tuple

try:
    from googleapiclient.discovery import build
    from google.oauth2.credentials import Credentials
    GOOGLE_API_AVAILABLE = True
except ImportError:
    GOOGLE_API_AVAILABLE = False


RATING_OPTIONS = ["Excellent", "Very Good", "Good", "Fair", "Poor"]

FORMS_SCOPE = "https://www.googleapis.com/auth/forms.body"


def get_oauth_config() -> Dict:
    """Get Google OAuth configuration from Streamlit secrets or environment."""
    try:
        import streamlit as st
        oauth_cfg = st.secrets.get("google_oauth", {})
        return {
            "client_id": oauth_cfg.get("client_id", ""),
            "client_secret": oauth_cfg.get("client_secret", ""),
            "redirect_uri": oauth_cfg.get(
                "redirect_uri",
                "http://localhost:8501/component/streamlit_oauth.authorize_button"
            ),
        }
    except Exception:
        return {
            "client_id": os.environ.get("GOOGLE_CLIENT_ID", ""),
            "client_secret": os.environ.get("GOOGLE_CLIENT_SECRET", ""),
            "redirect_uri": os.environ.get(
                "GOOGLE_REDIRECT_URI",
                "http://localhost:8501/component/streamlit_oauth.authorize_button"
            ),
        }


def build_forms_service(token: str):
    """Build a Google Forms API service client from an access token."""
    if not GOOGLE_API_AVAILABLE:
        raise ImportError(
            "google-api-python-client not installed. "
            "Run: pip install google-api-python-client google-auth"
        )

    credentials = Credentials(token=token)
    service = build("forms", "v1", credentials=credentials, cache_discovery=False)
    return service


def _make_short_answer_request(title: str, required: bool) -> Dict:
    """Create a batchUpdate request item for a short answer question."""
    return {
        "createItem": {
            "item": {
                "title": title,
                "questionItem": {
                    "question": {
                        "required": required,
                        "textQuestion": {
                            "paragraph": False,
                        },
                    }
                },
            },
            "location": {"index": 0},
        }
    }


def _make_paragraph_request(title: str, required: bool) -> Dict:
    """Create a batchUpdate request item for a paragraph question."""
    return {
        "createItem": {
            "item": {
                "title": title,
                "questionItem": {
                    "question": {
                        "required": required,
                        "textQuestion": {
                            "paragraph": True,
                        },
                    }
                },
            },
            "location": {"index": 0},
        }
    }


def _make_multiple_choice_request(title: str, options: List[str], required: bool) -> Dict:
    """Create a batchUpdate request item for a multiple choice question."""
    return {
        "createItem": {
            "item": {
                "title": title,
                "questionItem": {
                    "question": {
                        "required": required,
                        "choiceQuestion": {
                            "type": "RADIO",
                            "options": [{"value": opt} for opt in options],
                            "shuffle": False,
                        },
                    }
                },
            },
            "location": {"index": 0},
        }
    }


def create_google_form(
    access_token: str,
    survey_rows: List[Dict],
    form_title: str = "Workshop Feedback Survey",
) -> Tuple[str, str]:
    """
    Create a Google Form from the survey question list using the Forms API.

    Args:
        access_token: OAuth 2.0 access token with forms.body scope.
        survey_rows: Full survey from build_full_survey().
        form_title: Title for the Google Form.

    Returns:
        Tuple of (form_edit_url, form_published_url)
    """
    service = build_forms_service(access_token)

    # Step 1: Create the form with just a title
    form_body = {
        "info": {
            "title": form_title,
            "documentTitle": form_title,
        }
    }
    created_form = service.forms().create(body=form_body).execute()
    form_id = created_form["formId"]

    # Step 2: Build batchUpdate requests (reverse order — API inserts at index 0)
    requests = []
    for row in reversed(survey_rows):
        qtype = row.get("question_type", "Short Answer")
        question = str(row.get("question", ""))
        required = row.get("required", "No") == "Yes"
        options_str = row.get("options", "")
        options = [o.strip() for o in options_str.split("|") if o.strip()] if options_str else RATING_OPTIONS

        if qtype == "Short Answer":
            requests.append(_make_short_answer_request(question, required))
        elif qtype == "Paragraph":
            requests.append(_make_paragraph_request(question, required))
        elif qtype == "Multiple Choice":
            requests.append(_make_multiple_choice_request(question, options, required))

    # Step 3: Execute batchUpdate to add all questions
    if requests:
        service.forms().batchUpdate(
            formId=form_id,
            body={"requests": requests},
        ).execute()

    edit_url = f"https://docs.google.com/forms/d/{form_id}/edit"
    published_url = f"https://docs.google.com/forms/d/{form_id}/viewform"

    return edit_url, published_url
