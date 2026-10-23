# 📋 AI Survey Generator

An AI-powered Streamlit application that turns a workshop/training agenda into a complete, editable feedback survey. Uses Google Gemini to understand the agenda, generates session-specific feedback questions, and lets you export to Excel or create a Google Form directly.

---

## ✨ Features

- 📄 Upload a workshop agenda in **DOCX** format (or enter text manually)
- 📊 Optionally upload an existing feedback form (**XLSX**) as a style/reference
- 🤖 Uses **Gemini AI** to identify actual agenda sessions and generate one specific feedback question per session
- 🗂️ Automatically adds participant fields (Full Name, Designation, Organisation)
- ⭐ Automatically adds overall rating and suggestions questions
- ✏️ Review, edit, add, and delete questions in a visual data editor
- ⬇️ Export to **Excel** with styled formatting
- 📜 Download a **Google Apps Script** as a fallback workflow
- 🔗 Connect your **Google account via OAuth 2.0** and create a Google Form directly

---

## 🚀 Quick Start

### 1. Clone and set up

```bash
git clone https://github.com/your-username/ai-survey-generator.git
cd ai-survey-generator
python -m venv venv
# Windows:
venv\Scripts\activate
# macOS/Linux:
source venv/bin/activate
pip install -r requirements.txt
```

### 2. Configure Gemini API Key

Create a `.env` file (copy from `.env.example`):

```
GEMINI_API_KEY=your_gemini_api_key_here
```

Get your key at: https://aistudio.google.com/app/apikey

### 3. Run the app

```bash
streamlit run app.py
```

Open: http://localhost:8501

---

## 📝 Survey Structure

| Section | Questions |
|---------|-----------|
| Participant Information | Full Name, Designation, Organisation/Department |
| Session Feedback | One multiple-choice question per agenda session |
| Overall Workshop | Overall workshop rating |
| Suggestions | Open-ended improvement suggestions |

**Rating scale:** Excellent → Very Good → Good → Fair → Poor

---

## 📊 Excel Export Columns

| Column | Description |
|--------|-------------|
| No. | Sequential question number |
| Section | Survey section |
| Session No. | Workshop session number |
| Session | Session title |
| Category | Question category |
| Question | Feedback question text |
| Question Type | Short Answer / Multiple Choice / Paragraph |
| Options | Answer choices (pipe-separated) |
| Required | Yes or No |

---

## 🔗 Google Forms Integration

### Option 1: Apps Script (Recommended — no setup needed)
1. Download the `.gs` file from the Export step
2. Go to [script.google.com](https://script.google.com)
3. Create New Project → Paste script → Run `createSurveyForm`
4. Authorize and the form is created in your account

### Option 2: Direct OAuth Integration
Configure `.streamlit/secrets.toml` (see `secrets.example.toml`):

```toml
[google_oauth]
client_id = "YOUR_CLIENT_ID.apps.googleusercontent.com"
client_secret = "YOUR_CLIENT_SECRET"
redirect_uri = "http://localhost:8501/component/streamlit_oauth.authorize_button"
```

Required: Google Forms API enabled + OAuth credentials in Google Cloud Console.

---

## 🛠️ Technologies

| Library | Purpose |
|---------|---------|
| `streamlit` | Web application UI |
| `google-generativeai` | Gemini AI content generation |
| `python-docx` | DOCX agenda parsing |
| `openpyxl` | Excel export |
| `streamlit-oauth` | Google OAuth 2.0 flow |
| `google-api-python-client` | Google Forms API |
| `python-dotenv` | Local environment variables |

---

## 🔒 Security Notes

- Never commit `.env` or `.streamlit/secrets.toml` to version control
- Use Streamlit Cloud Secrets for production deployments
- OAuth tokens are stored in session state only (not persisted)

---

## ☁️ Streamlit Cloud Deployment

1. Push to GitHub (`.env` and `secrets.toml` are gitignored)
2. Connect repo to [Streamlit Community Cloud](https://streamlit.io/cloud)
3. Set `app.py` as main file
4. Add secrets in the Streamlit app settings
5. Register deployed callback URI in Google Cloud Console:
   ```
   https://YOUR-APP.streamlit.app/component/streamlit_oauth.authorize_button
   ```
