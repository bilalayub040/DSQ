import os
import sys
import glob
import shutil
import datetime
import win32com.client
import tkinter as tk
from tkinter import ttk

# --- Determine base directory properly for .exe ---
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def find_today_files(folder):
    today_str = datetime.datetime.now().strftime("%#d%b%y")  # e.g., 5Oct25
    files = []
    for f in glob.glob(os.path.join(folder, "*_submit_*.txt")):
        basename = os.path.basename(f)
        name_part = os.path.splitext(basename)[0]  # remove .txt
        if name_part.rsplit("_", 1)[-1].lower() == today_str.lower():
            files.append(f)
    return files

def parse_file(file_path):
    subject = ""
    to_emails = ""
    cc_emails = ""
    attachments = []
    body_lines = []
    with open(file_path, "r", encoding="utf-8") as f:
        content = f.read()
    lines = content.splitlines()
    mode = None
    for line in lines:
        if line.startswith("Subject:"):
            subject = line.replace("Subject:", "").strip()
        elif line.startswith("To:"):
            to_emails = line.replace("To:", "").strip()
        elif line.startswith("Cc:"):
            cc_emails = line.replace("Cc:", "").strip()
        elif line.startswith("Attachments:"):
            mode = "attachments"
        elif line.startswith("Body:"):
            mode = "body"
        else:
            if mode == "attachments" and line.strip():
                attachments.append(line.strip())
            elif mode == "body":
                body_lines.append(line)
    body = "\n".join(body_lines)
    return subject, to_emails, cc_emails, attachments, body

def send_email(subject, to_emails, cc_emails, attachments, body, status_window):
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        # Get the default sending account email
        sender_email = namespace.Accounts.Item(1).SmtpAddress

        mail = outlook.CreateItem(0)  # 0=MailItem
        mail.Subject = subject
        mail.To = to_emails.replace(";", ",")
        
        # Build CC: existing + sender email
        cc_list = [e.strip() for e in cc_emails.split(";") if e.strip()]
        if sender_email not in cc_list:
            cc_list.append(sender_email)
        mail.CC = ", ".join(cc_list)
        
        mail.HTMLBody = body

        # Add attachments
        for att in attachments:
            if os.path.exists(att):
                mail.Attachments.Add(att)

        status_window.status_label.config(text="Status: Sending...")
        mail.Send()
        status_window.status_label.config(text="Status: Sent successfully!")
        status_window.after(3000, status_window.destroy)
    except Exception as e:
        status_window.status_label.config(text=f"Status: Failed\n{e}")
        status_window.after(5000, status_window.destroy)

class StatusWindow(tk.Tk):
    def __init__(self, to_emails, cc_emails, subject):
        super().__init__()
        self.title("Email Sending Status")
        self.geometry("400x120+500+300")
        self.resizable(False, False)
        self.attributes("-topmost", True)
        # Make draggable
        self.bind("<ButtonPress-1>", self.start_move)
        self.bind("<ButtonRelease-1>", self.stop_move)
        self.bind("<B1-Motion>", self.do_move)
        self.offset_x = 0
        self.offset_y = 0

        ttk.Label(self, text=f"Sending to: {to_emails} and {cc_emails}", wraplength=380).pack(pady=5)
        ttk.Label(self, text=f"Subject: {subject}", wraplength=380).pack(pady=5)
        self.status_label = ttk.Label(self, text="Status: Pending", wraplength=380)
        self.status_label.pack(pady=5)

    def start_move(self, event):
        self.offset_x = event.x
        self.offset_y = event.y

    def stop_move(self, event):
        self.offset_x = 0
        self.offset_y = 0

    def do_move(self, event):
        x = self.winfo_pointerx() - self.offset_x
        y = self.winfo_pointery() - self.offset_y
        self.geometry(f"+{x}+{y}")

def main():
    folder = BASE_DIR
    files = find_today_files(folder)
    if not files:
        print("No files to send today.")
        return

    for fpath in files:
        subject, to_emails, cc_emails, attachments, body = parse_file(fpath)
        status_window = StatusWindow(to_emails, cc_emails, subject)
        status_window.after(100, lambda: send_email(subject, to_emails, cc_emails, attachments, body, status_window))
        status_window.mainloop()

    # Delete all files including itself and folder
    try:
        for f in os.listdir(folder):
            file_path = os.path.join(folder, f)
            if os.path.isfile(file_path):
                os.remove(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        # Delete folder itself
        os.chdir(os.path.dirname(folder))
        shutil.rmtree(folder)
    except Exception as e:
        print(f"Cleanup failed: {e}")

if __name__ == "__main__":
    main()


