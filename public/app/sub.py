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

        # --- Find DSQ.qa account ---
        dsq_account = None
        for acc in namespace.Accounts:
            if "dsq.qa" in acc.SmtpAddress.lower():
                dsq_account = acc
                break

        if not dsq_account:
            raise Exception("No DSQ.qa account found in Outlook!")

        sender_email = dsq_account.SmtpAddress.lower()

        # --- Create the mail inside DSQ account’s store ---
        draft_folder = dsq_account.DeliveryStore.GetDefaultFolder(16)  # olFolderDrafts = 16
        mail = outlook.CreateItem(0)
        mail.Move(draft_folder)  # move mail into DSQ Drafts (now belongs to that account)

        # --- Clean recipients ---
        to_list = [e.strip() for e in to_emails.replace(",", ";").split(";") if e.strip()]
        cc_list = [e.strip() for e in cc_emails.replace(",", ";").split(";") if e.strip()]

        mail.Subject = subject
        mail.To = "; ".join(to_list)
        mail.CC = "; ".join(cc_list)
        mail.HTMLBody = body

        # --- Attachments ---
        for att in attachments:
            if os.path.exists(att):
                mail.Attachments.Add(att)

        # --- Ensure Sent Items of same account ---
        sent_folder = dsq_account.DeliveryStore.GetDefaultFolder(5)
        mail.SaveSentMessageFolder = sent_folder

        # --- Send email ---
        status_window.status_label.config(text=f"Status: Sending via {sender_email}...")
        mail.Send()
        status_window.status_label.config(text="Status: Sent successfully!")
        status_window.after(3000, status_window.destroy)

    except Exception as e:
        status_window.status_label.config(text=f"Status: Failed\n{e}")
        status_window.after(6000, status_window.destroy)

class StatusWindow(tk.Tk):
    def __init__(self, to_emails, cc_emails, subject):
        super().__init__()
        self.title("Email Sending Status")
        self.geometry("400x120+500+300")
        self.resizable(False, False)
        self.attributes("-topmost", True)

        # --- Make draggable ---
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
        status_window.after(
            100,
            lambda s=subject, t=to_emails, c=cc_emails, a=attachments, b=body, w=status_window:
            send_email(s, t, c, a, b, w)
        )
        status_window.mainloop()

    # --- Cleanup processed files and itself ---
    try:
        for f in files:
            if os.path.exists(f) and os.path.isfile(f):
                os.remove(f)

        script_path = sys.executable if getattr(sys, 'frozen', False) else __file__
        os.remove(script_path)
    except Exception as e:
        print(f"Cleanup failed: {e}")

if __name__ == "__main__":
    main()

