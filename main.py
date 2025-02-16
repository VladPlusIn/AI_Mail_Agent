import tkinter as tk
from tkinter import messagebox, ttk
import subprocess
import sys
import os
import json
import pandas as pd
import win32com.client
from datetime import datetime, timedelta
import pytz

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(BASE_DIR, "config.json")
SCRIPT_PATH = os.path.join(BASE_DIR, "email_ai_app.py")
LOG_JSON_FILE = os.path.join(BASE_DIR, "email_rag_log.jsonl")

def run_application():
    """Runs the main script and then displays the log table."""
    try:
        subprocess.run([sys.executable, SCRIPT_PATH], check=True)
        messagebox.showinfo("Success", "Application executed successfully!")
        # Slight delay before loading the log table (to ensure the log file is updated).
        root.after(1000, display_log_table)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to run the application:\n{e}")

def setup_application():
    """Initial setup for different users via GUI instead of console input."""
    setup_window = tk.Toplevel(root)
    setup_window.title("Setup Application")

    tk.Label(setup_window, text="Enter OpenAI API Key:").pack()
    api_key_entry = tk.Entry(setup_window, width=50)
    api_key_entry.pack()

    tk.Label(setup_window, text="Enter Your Email:").pack()
    user_email_entry = tk.Entry(setup_window, width=50)
    user_email_entry.pack()

    tk.Label(setup_window, text="Enter AI Response Prompt:").pack()
    ai_response_prompt_entry = tk.Entry(setup_window, width=50)
    ai_response_prompt_entry.pack()

    # Additional config fields for days & classification prompts
    tk.Label(setup_window, text="Number of days for unread emails (default=3):").pack()
    days_for_unread_entry = tk.Entry(setup_window, width=50)
    days_for_unread_entry.insert(0, "3")
    days_for_unread_entry.pack()

    tk.Label(setup_window, text="Prompt for NEED to Reply (default='directly addressed...'):").pack()
    prompt_need_entry = tk.Entry(setup_window, width=50)
    prompt_need_entry.insert(0, "directly addressed, question, or complaint")
    prompt_need_entry.pack()

    tk.Label(setup_window, text="Prompt for MIGHT Reply (default='general request...'):").pack()
    prompt_might_entry = tk.Entry(setup_window, width=50)
    prompt_might_entry.insert(0, "general request where user is in CC")
    prompt_might_entry.pack()

    tk.Label(setup_window, text="Prompt for MAY NOT Reply (default='no response needed'):").pack()
    prompt_maynot_entry = tk.Entry(setup_window, width=50)
    prompt_maynot_entry.insert(0, "no response needed")
    prompt_maynot_entry.pack()

    def save_config():
        config_data = {
            "OPENAI_API_KEY": api_key_entry.get(),
            "USER_EMAIL": user_email_entry.get(),
            "AI_RESPONSE_PROMPT": ai_response_prompt_entry.get(),
            "DAYS_FOR_UNREAD_EMAIL": days_for_unread_entry.get(),
            "PROMPT_NEED_REPLY": prompt_need_entry.get(),
            "PROMPT_MIGHT_REPLY": prompt_might_entry.get(),
            "PROMPT_MAYNOT_REPLY": prompt_maynot_entry.get()
        }
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config_data, f, indent=4)
        messagebox.showinfo("Success", "Configuration saved successfully!")
        setup_window.destroy()

    save_button = tk.Button(setup_window, text="Save", command=save_config)
    save_button.pack()

def display_log_table():
    """Displays the email log table from JSON log."""
    if not os.path.exists(LOG_JSON_FILE):
        messagebox.showerror("Error", "Log file not found!")
        return

    log_window = tk.Toplevel(root)
    log_window.title("Email Log Table")

    with open(LOG_JSON_FILE, "r", encoding="utf-8") as f:
        log_entries = [json.loads(line) for line in f if line.strip()]
        # Sort logs by importance: Need -> Might -> May Not
        order = ['Need to Reply', 'Might Reply', 'May Not Reply']
        log_entries = sorted(
            log_entries,
            key=lambda x: order.index(x.get('importance', 'May Not Reply'))
        )
    
    if not log_entries:
        messagebox.showinfo("Info", "No log data available.")
        return
    
    columns = ['timestamp', 'subject', 'sender', 'received_time', 'body_summary', 'ai_response', 'importance']
    tree = ttk.Treeview(log_window, columns=columns, show='headings', height=20)
    
    for col in columns:
        tree.heading(col, text=col)
        tree.column(col, anchor="w", minwidth=150, width=250, stretch=True)
    
    for entry in log_entries:
        tree.insert("", "end", values=[entry.get(col, "N/A") for col in columns])
    
    tree.pack(expand=True, fill='both')

    # Add tooltip functionality to show full text on hover
    def on_hover(event):
        item = tree.identify_row(event.y)
        column = tree.identify_column(event.x)
        if item and column:
            col_index = int(column[1:]) - 1
            values = tree.item(item, 'values')
            if col_index < len(values):
                tooltip_text.set(values[col_index])
                tooltip.place(x=event.x_root, y=event.y_root)
    
    def on_leave(event):
        tooltip.place_forget()
    
    tooltip_text = tk.StringVar()
    tooltip = tk.Label(log_window, textvariable=tooltip_text, background='yellow', wraplength=400)
    
    tree.bind('<Motion>', on_hover)
    tree.bind('<Leave>', on_leave)

def check_setup():
    """Ensure setup_application() runs when config.json is missing."""
    if not os.path.exists(CONFIG_FILE):
        setup_application()

# GUI for Main Window
root = tk.Tk()
root.title("Email AI Assistant")

frame = tk.Frame(root, padx=20, pady=20)
frame.pack(pady=10)

tk.Label(frame, text="Welcome to AI Email Assistant").pack()

run_button = tk.Button(frame, text="Run Application", command=run_application)
run_button.pack(pady=5)

setup_button = tk.Button(frame, text="Setup Application", command=setup_application)
setup_button.pack(pady=5)

log_button = tk.Button(frame, text="View Log Table", command=display_log_table)
log_button.pack(pady=5)

exit_button = tk.Button(frame, text="Exit", command=root.quit)
exit_button.pack(pady=5)

check_setup()
root.mainloop()






