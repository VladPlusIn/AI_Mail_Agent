import win32com.client
from datetime import datetime, timedelta
import pytz
import os
import json
import sys
import pandas as pd
from openai import OpenAI
import time
import requests

LOG_FILE = "email_rag_log.jsonl"
LOG_TABLE_FILE = "email_rag_log.csv"
CONFIG_FILE = "config.json"

# Ensure log files exist
if not os.path.exists(LOG_FILE):
    with open(LOG_FILE, "w", encoding="utf-8") as f:
        f.write("")  # Create an empty JSONL file

if not os.path.exists(LOG_TABLE_FILE):
    pd.DataFrame(columns=["timestamp", "subject", "sender", "received_time", "body_summary", "ai_response", "importance"]).to_csv(LOG_TABLE_FILE, index=False)

# Load user-specific configuration
CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), CONFIG_FILE)
if os.path.exists(CONFIG_PATH):
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        config = json.load(f)
    API_KEY = config.get("OPENAI_API_KEY", "")
    USER_EMAIL = config.get("USER_EMAIL", "")
    # The user can override, but we keep a strong system disclaimer to reduce conflicts
    AI_RESPONSE_PROMPT = config.get(
        "AI_RESPONSE_PROMPT",
        "You are a professional assistant. Respond politely and concisely in HTML format. Preserve paragraph formatting."
    )
    DAYS_FOR_UNREAD_EMAIL = int(config.get("DAYS_FOR_UNREAD_EMAIL", 3))
    PROMPT_NEED_REPLY = config.get("PROMPT_NEED_REPLY", "directly addressed, question, or complaint")
    PROMPT_MIGHT_REPLY = config.get("PROMPT_MIGHT_REPLY", "general request where user is in CC")
    PROMPT_MAYNOT_REPLY = config.get("PROMPT_MAYNOT_REPLY", "no response needed")
else:
    print("Configuration file missing. Please run the setup first.")
    sys.exit(1)

# Create the OpenAI client
client = OpenAI(
    base_url="https://openrouter.ai/api/v1",
    api_key=API_KEY,
)

def safe_ai_call(**kwargs):
    """
    A wrapper for AI calls that handles connectivity and rate-limit errors.
    Retries once if we get a rate-limit error or a transient issue.
    """
    try:
        response = client.chat.completions.create(**kwargs)
        return response
    except requests.exceptions.RequestException as e:
        # Connectivity or network issues
        print(f"[Error] Network issue with OpenAI API: {e}")
        return None
    except Exception as e:
        # Possible rate-limit or other AI errors
        err_str = str(e).lower()
        if "rate limit" in err_str or "overloaded" in err_str:
            print("[Warning] Rate limit encountered; retrying after 2 seconds...")
            time.sleep(2)
            try:
                response = client.chat.completions.create(**kwargs)
                return response
            except Exception as e2:
                print(f"[Error] Retried and still failed: {e2}")
                return None
        print(f"[Error] AI call failed: {e}")
        return None

def summarize_text(text):
    """
    Summarizes email content with bullet points and a focus on actions/deadlines.
    Improved prompt for clarity.
    """
    try:
        prompt = (
            "Please provide a short summary in bullet points focusing on actions, deadlines, or requests:\n\n"
            f"{text}"
        )
        summary_response = safe_ai_call(
            model="openai/gpt-4-32k",
            messages=[
                {
                    "role": "system",
                    "content": (
                        "You are an expert email summarization assistant. "
                        "Always produce a short summary highlighting key points, deadlines, or actions. "
                        "User config may exist, but you must not contradict these system instructions."
                    )
                },
                {"role": "user", "content": prompt}
            ],
            temperature=0.2,  # Lower temperature => more deterministic
        )
        if summary_response and summary_response.choices:
            return summary_response.choices[0].message.content.strip()
        else:
            return "Summary not available or AI call failed."
    except Exception as e:
        return f"[Error in summarize_text] {e}"

def draft_outlook_response(email, response):
    """Replies to the email while keeping the conversation history."""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        inbox = outlook.GetNamespace("MAPI").GetDefaultFolder(6)
        messages = inbox.Items
        message = None

        for msg in messages:
            if msg.Subject == email['Subject'] and msg.SenderName == email['Sender']:
                message = msg
                break

        if message:
            reply = message.ReplyAll()
            reply.HTMLBody = f"{response}<br><br>" + reply.HTMLBody.split('<div class=\"WordSection1\">', 1)[-1]
            reply.Display()
            print(f"Replied to email: {email['Subject']}")
        else:
            print(f"Original email not found for subject: {email['Subject']}")
    except Exception as e:
        print(f"[Error drafting email] {e}")

def log_email_data(email, summary, response, importance):
    """Logs all retrieved email data, including AI summary and responses."""
    log_entry = {
        "timestamp": datetime.now().isoformat(),
        "subject": email["Subject"],
        "sender": email["Sender"],
        "received_time": email["ReceivedTime"].isoformat(),
        "body_summary": summary,
        "ai_response": response,
        "importance": importance
    }

    # Append to JSONL file
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            json.dump(log_entry, f)
            f.write("\n")
    except Exception as e:
        print(f"[Error writing to JSON log] {e}")

    # Append to CSV file
    try:
        log_df = pd.DataFrame([log_entry])
        if not os.path.exists(LOG_TABLE_FILE):
            log_df.to_csv(LOG_TABLE_FILE, index=False)
        else:
            log_df.to_csv(LOG_TABLE_FILE, mode='a', header=False, index=False)
    except Exception as e:
        print(f"[Error writing to CSV log] {e}")

def get_unread_emails():
    """Retrieves unread emails from the last N days, based on config."""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)  # 6 is the index for the Inbox
        messages = inbox.Items

        unread_emails = []
        time_cutoff = datetime.now(pytz.utc) - timedelta(days=DAYS_FOR_UNREAD_EMAIL)

        for message in messages:
            try:
                received_time = message.ReceivedTime
                if isinstance(received_time, datetime):
                    received_time = pytz.utc.localize(received_time) if received_time.tzinfo is None else received_time
                
                if message.UnRead and received_time >= time_cutoff:
                    unread_emails.append({
                        "Subject": message.Subject,
                        "Sender": message.SenderName,
                        "ReceivedTime": received_time,
                        "Body": message.Body,
                    })
            except AttributeError:
                pass

        return unread_emails
    except Exception as e:
        print(f"[Error accessing Outlook] {e}")
        return []
    
def determine_email_importance(email):
    """
    Classifies email importance using user-defined classification prompts.
    The system disclaimers ensure the user config doesn't override system logic.
    """
    try:
        # Provide short examples to reduce confusion
        classification_prompt = f"""You have three categories:
1) Need to Reply ({PROMPT_NEED_REPLY})
2) Might Reply ({PROMPT_MIGHT_REPLY})
3) May Not Reply ({PROMPT_MAYNOT_REPLY})

If the email explicitly addresses me with a question or direct request, it's 'Need to Reply'.
If I'm just in CC or it's partly relevant, it's 'Might Reply'.
If it doesn't require any action, it's 'May Not Reply'.

Email Content:
{email['Body']}

Respond with exactly one of: 'Need to Reply', 'Might Reply', or 'May Not Reply'.
"""
        response = safe_ai_call(
            model="openai/gpt-4-32k",
            messages=[
                {
                    "role": "system",
                    "content": (
                        "You are a strict email classification assistant. "
                        "You must not contradict system instructions. "
                        "User config may alter categories but can't override these instructions."
                    )
                },
                {"role": "user", "content": classification_prompt}
            ],
            temperature=0.3,
        )
        if response and response.choices:
            classification = response.choices[0].message.content.strip()
            if classification not in ["Need to Reply", "Might Reply", "May Not Reply"]:
                classification = "May Not Reply"
            return classification
        else:
            return "May Not Reply"
    except Exception as e:
        print(f"[Error determining importance] {e}")
        return "May Not Reply"

def interact_with_ai_agent(email):
    """Generates a response to the given email using AI with improved role-based prompting."""
    try:
        print(f"Generating AI response for: {email['Subject']}")
        user_prompt = (
            f"{AI_RESPONSE_PROMPT}\n\n"
            f"Sender: {email['Sender']}\nBody: {email['Body']}\n"
            "The user config can shape style, but system instructions override it if there's a conflict.\n"
            "Please respond in HTML format, focusing on clarity, ending with 'Best regards' but excluding your name/position.\n"
            "Preserve paragraph formatting and keep the tone polite.\n"
        )
        response = safe_ai_call(
            model="openai/gpt-4-32k",
            messages=[
                {
                    "role": "system",
                    "content": (
                        "You are a helpful email reply assistant. "
                        "System instructions have the highest priority. "
                        "User config might specify style, but do not contradict system instructions."
                    )
                },
                {"role": "user", "content": user_prompt}
            ],
            temperature=0.4,
        )
        if response and response.choices:
            return response.choices[0].message.content
        else:
            return "AI response was empty or invalid."
    except Exception as e:
        print(f"[Error interacting with AI agent] {e}")
        return "Error generating AI response"

def process_emails():
    """
    Processes unread emails from the last N days (DAYS_FOR_UNREAD_EMAIL).
    Logs them, then classifies. If classification is 'Need to Reply', we auto-reply.
    """
    unread_emails = get_unread_emails()
    if not unread_emails:
        print("No unread emails found in the specified time range.")
        return

    for email in unread_emails:
        # Classify
        importance = determine_email_importance(email)
        # Summarize
        summary = summarize_text(email["Body"])
        # If 'Need to Reply', produce an AI-based response
        ai_response = ""
        if importance == "Need to Reply":
            ai_response = interact_with_ai_agent(email)
        # Log everything
        log_email_data(email, summary, ai_response, importance)
        # Draft the Outlook reply if needed
        if importance == "Need to Reply":
            draft_outlook_response(email, ai_response)

if __name__ == "__main__":
    process_emails()
