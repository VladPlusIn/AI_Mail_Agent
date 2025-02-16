# Email AI Assistant

This project provides an **AI-powered email processing tool** that integrates with Microsoft Outlook to:

1. **Fetch unread emails** within a user-specified timeframe.
2. **Summarize and classify** each email using GPT-based AI.
3. **Draft Outlook responses** for emails categorized as “Need to Reply.”

## 📂 Repository Structure

- **`main.py`** – A Tkinter-based GUI for configuration and running the assistant.
- **`email_ai_app.py`** – The core logic that interacts with Outlook, the OpenAI API, and logs results.
- **`config.json`** – A file storing user credentials and preferences.
- **`requirements.txt`** – List of dependencies for pip installation.

---

## 📌 Features

✔ **GUI Setup**: Configure API keys, email address, number of days to check for unread emails, and classification prompts.  
✔ **Automated Classification**:
   - **Need to Reply** – (Default: "directly addressed, question, or complaint")
   - **Might Reply** – (Default: "general request where user is in CC")
   - **May Not Reply** – (Default: "no response needed")  
✔ **Auto-Drafting in Outlook**: Creates email drafts for "Need to Reply" messages.  
✔ **Logging**: Summaries, classification, and responses are saved in:
   - `email_rag_log.jsonl` (JSON Lines)
   - `email_rag_log.csv`  
✔ **Tooltip for Log Viewing**: Hover over table entries to view full text in a tooltip.

---

## 🔧 Prerequisites

1. **Windows** system with **Microsoft Outlook** installed and configured.
2. **Python 3.7+** installed.
3. Install dependencies:
   ```bash
   pip install -r requirements.txt

---

## ⚙ Configuration & Setup
A sample `config.json` file is provided.

Configuration Parameters:


| Parameter               | Description |
|-------------------------|-------------|
| `OPENAI_API_KEY`       | Your API key from OpenAI/OpenRouter. |
| `USER_EMAIL`           | Your primary email address. |
| `AI_RESPONSE_PROMPT`   | AI prompt that defines the style and tone of email responses. |
| `DAYS_FOR_UNREAD_EMAIL` | Number of days to check unread emails (default: `3`). |
| `PROMPT_NEED_REPLY`    | Custom classification instruction for "Need to Reply" emails. |
| `PROMPT_MIGHT_REPLY`   | Custom classification instruction for "Might Reply" emails. |
| `PROMPT_MAYNOT_REPLY`  | Custom classification instruction for "May Not Reply" emails. |


If config.json is missing, the GUI will prompt the user to configure settings.


## 📌 Usage

## 1️⃣ Run `main.py`

python main.py

* A Tkinter window will appear.

### 🔹 2️⃣ Click **"Setup Application"** 

* Configure:
  * **API Key**
  * **Email Address**
  * **AI Prompt**
  * **Number of days for** email retrieval
  * **Classification prompts**

### 🔹 3️⃣ Click **"Run Application"**

* **Fetches** unread emails from the last `N` days.
* **Classifies** emails based on AI-generated responses.
* **Logs results** and **drafts Outlook replies**.

### 🔹 4️⃣ Click **"View Log Table"**

* **Opens a table** displaying processed emails.
* **Hover over cells** to see **tooltips with full text**.

## 🛠 Troubleshooting

### ❌ Missing `config.json`

* The script will **prompt a Setup GUI** if it cannot find `config.json`.

### ❌ Outlook Not Installed or Configured

* `win32com.client` calls may fail if Outlook isn’t installed properly.

### ❌ API Rate Limits

* If you see `[Warning] Rate limit encountered...`, the script **automatically retries** if overloaded.

### ❌ No Emails Are Found

* Check `DAYS_FOR_UNREAD_EMAIL` in `config.json` (set it to `14` for testing).
* Ensure you have unread emails in your Outlook **Inbox**.

### ❌ Prompt Conflicts

* AI instructions are **system-anchored**. If user-supplied prompts are **contradictory**, the system may override them.

### ❌ Permission Issues

* If Windows blocks the `.exe`, try **running as Administrator**.

---

## 📜 License

This project is under the **MIT License**. See `LICENSE` file for details.

---

## 🙌 Contributions

* **Maintainers**: Vlad Plyusnin
* **Pull Requests** are welcome! Open an issue before submitting major changes.

