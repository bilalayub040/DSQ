import os
import sys
import glob
import datetime
import re
import win32com.client
import pythoncom
import tkinter as tk
from tkinter import ttk
import threading

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
        name_part = os.path.splitext(basename)[0]
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
    def ui_set(text):
        try:
            status_window.after(0, lambda: status_window.status_label.config(text=text))
        except Exception:
            pass

    def ui_close(delay_ms=3000):
        try:
            status_window.after(delay_ms, status_window.destroy)
        except Exception:
            pass

    def worker():
        # initialize COM for this thread
        pythoncom.CoInitialize()
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            session = outlook.Session
            namespace = outlook.GetNamespace("MAPI")

            # ----- Resolve sender email with multiple fallbacks -----
            sender_email = ""
            try:
                # best attempt for Exchange users
                ae = session.CurrentUser.AddressEntry
                if ae:
                    ex = ae.GetExchangeUser()
                    if ex and ex.PrimarySmtpAddress:
                        sender_email = ex.PrimarySmtpAddress
            except Exception:
                sender_email = ""

            if not sender_email:
                # fallback to first account SMTP if possible
                try:
                    acc = namespace.Accounts.Item(1)
                    if acc and getattr(acc, "SmtpAddress", None):
                        sender_email = acc.SmtpAddress
                except Exception:
                    sender_email = ""

            sender_email = (sender_email or "").strip().lower()

            # ----- Parse recipients (accept ',' or ';') -----
            def split_addrs(s):
                if not s:
                    return []
                parts = [p.strip() for p in re.split(r'[;,]', s) if p and p.strip()]
                return parts

            to_list = split_addrs(to_emails)
            cc_list = split_addrs(cc_emails)

            # ----- Add sender to CC if resolved and not already present -----
            if sender_email:
                lower_cc = [c.lower() for c in cc_list]
                if sender_email not in lower_cc:
                    cc_list.append(sender_email)

            # Build strings shown to Outlook
            to_str = "; ".join(to_list)
            cc_str = "; ".join(cc_list)

            # update UI to show final To/CC
            ui_set(f"Sending to: {to_str or '(none)'}\nCC: {cc_str or '(none)'}\nStatus: Sending...")

            if not to_list:
                ui_set("Status: Failed\nNo valid 'To' email address found!")
                ui_close(5000)
                return

            # ----- Create and fill MailItem -----
            mail = outlook.CreateItem(0)
            mail.Subject = subject or ""
            mail.To = to_str
            mail.CC = cc_str

            # HTML body with fallback
            try:
                mail.HTMLBody = body or ""
            except Exception:
                try:
                    mail.Body = re.sub(r'<[^>]+>', '', body or "")
                except Exception:
                    mail.Body = body or ""

            # Attachments (skip missing/bad)
            for att in attachments or []:
                try:
                    if att and os.path.exists(att):
                        mail.Attachments.Add(att)
                except Exception:
                    continue

            # Try to set SaveSentMessageFolder to sender's account Sent if possible (best-effort)
            try:
                target_acc = None
                try:
                    for acc in namespace.Accounts:
                        try:
                            if getattr(acc, "SmtpAddress", "").lower() == sender_email:
                                target_acc = acc
                                break
                        except Exception:
                            continue
                except Exception:
                    target_acc = None

                if target_acc:
                    try:
                        sent_folder = target_acc.DeliveryStore.GetDefaultFolder(5)  # olFolderSentMail
                        mail.SaveSentMessageFolder = sent_folder
                    except Exception:
                        pass
            except Exception:
                pass

            # Send
            mail.Send()
            ui_set("Status: Sent successfully!")
            ui_close(3000)

        except Exception as e:
            ui_set(f"Status: Failed\n{e}")
            ui_close(6000)
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    # run worker thread (daemon so it won't block exit)
    t = threading.Thread(target=worker, daemon=True)
    t.start()

class StatusWindow(tk.Tk):
    def __init__(self, to_emails, cc_emails, subject):
        super().__init__()
        self.title("Email Sending Status")
        self.geometry("420x120+500+300")
        self.resizable(False, False)
        self.attributes("-topmost", True)
        # Make draggable
        self.bind("<ButtonPress-1>", self.start_move)
        self.bind("<ButtonRelease-1>", self.stop_move)
        self.bind("<B1-Motion>", self.do_move)
        self.offset_x = 0
        self.offset_y = 0

        ttk.Label(self, text=f"Sending to: {to_emails or '(none)'}", wraplength=400).pack(pady=(8,0))
        ttk.Label(self, text=f"Cc: {cc_emails or '(none)'}", wraplength=400).pack(pady=(2,6))
        ttk.Label(self, text=f"Subject: {subject or ''}", wraplength=400).pack()
        self.status_label = ttk.Label(self, text="Status: Pending", wraplength=400)
        self.status_label.pack(pady=6)

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
        # schedule the send shortly after window appears
        status_window.after(100, lambda s=subject, t=to_emails, c=cc_emails, a=attachments, b=body, w=status_window:
                            send_email(s, t, c, a, b, w))
        status_window.mainloop()

    # --- Cleanup processed files and self (best-effort) ---
    try:
        for f in files:
            try:
                if os.path.exists(f) and os.path.isfile(f):
                    os.remove(f)
            except Exception:
                pass
        script_path = sys.executable if getattr(sys, 'frozen', False) else __file__
        try:
            os.remove(script_path)
        except Exception:
            # cannot delete running exe; ignore
            pass
    except Exception as e:
        print(f"Cleanup failed: {e}")

if __name__ == "__main__":
    main()

