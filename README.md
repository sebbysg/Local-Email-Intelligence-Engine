# 📧 Local Email Intelligence Engine (v1.1)

A local-first, privacy-focused IT Project Management assistant that parses Outlook emails using **Gemma 3 4B** (via Ollama) to generate daily action reports and "No-Reply" tracking.

## 🚀 Overview
This system automates the tedious task of morning email triage. It pulls data from your local Outlook application, processes it through a local LLM, and delivers a structured summary via Telegram while archiving a versioned audit trail in Microsoft Word format.

### Key Features
* **Local-First Privacy**: No email content ever leaves your machine (except for the encrypted Telegram notification).
* **Intelligent No-Reply Detection**: Automatically identifies outbound emails that haven't received a reply in 7 days.
* **Human Accountability**: Every summary point is tied to a **MAPI EntryID**, allowing for instant auditing of sources.
* **Sentiment Analysis**: Flags high-urgency or frustrated stakeholder tones.

---

## 🛠️ Project Structure
* `main.py`: The core engine for data extraction and LLM processing.
* `audit_tool.py`: A helper utility to pop open specific emails in Outlook using an EntryID.
* `templates/`: Contains customizable prompt personas (e.g., `pmo_standard.txt`).
* `reports/`: Local archive of generated Word (.docx) summaries.
* `.env`: (Excluded from Git) Stores your sensitive API tokens and local paths.

---

## 📦 Setup Instructions

### 1. Prerequisites
* Windows OS with **Outlook (Classic)** installed and configured.
* [Ollama](https://ollama.com/) running locally with `gemma3:4b`.
* Python 3.10+ and a virtual environment.

### 2. Installation

git clone [https://github.com/your-username/Email-Intelligence.git](https://github.com/your-username/Email-Intelligence.git)
cd Email-Intelligence
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt

### 3. Environment Variables
Create a .env file in the root directory:

TELEGRAM_TOKEN=your_bot_token
TELEGRAM_CHAT_ID=your_chat_id
OUTPUT_DIR=C:\Your\Path\To\Reports

### 4. Running the Tool
To run manually:
python main.py

To audit an email:
python audit_tool.py

🛡️ Values & Philosophy
This tool was built with the following guiding principles:

AI Assists, Humans Lead: AI is used to summarize, but the human remains responsible for verification via the Audit Trail.

People-First Tech: Automation should reduce burnout, not create more work.

Accessibility: Soft skills and clear communication are prioritized over pure technical complexity.

📄 License
This project is open-source. Please ensure you respect corporate data policies when using AI to process work-related emails.

