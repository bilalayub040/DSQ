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
    """
    Runs in a background thread. Initializes COM for the thread, sends the mail,
    updates UI through status_window.after(...) to stay thread-safe.
    """

    def ui_set(text):
        # schedule UI update on main thread
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
        pythoncom.CoInitialize()
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            session = outlook.Session

            # Get sender (default/current Outlook user) safely
            try:
                sender_email = session.CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress
            except Exception:
                # Fallback - display name (not ideal but safe)
                try:
                    sender_email = session.CurrentUser.AddressEntry.Name
                except Exception:
                    sender_email = ""
            sender_email = (sender_email or "").strip().lower()

            # Prepare recipients - accept commas or semicolons
            def split_addrs(s):
                if not s:
                    return []
                parts = re.split(r'[;,]', s)
                return [p.strip() for p in parts if p and p.strip()]

            to_list = split_addrs(to_emails)
            cc_list = split_addrs(cc_emails)

            # Add sender to CC if not already present (case-insensitive)
            if sender_email:
                lower_cc = [c.lower() for c in cc_list]
                if sender_email not in lower_cc:
                    cc_list.append(sender_email)

            # Validate at least one To
            if not to_list:
                ui_set("Status: Failed\nNo valid 'To' email address found!")
                ui_close(5000)
                return

            # Create mail item (use default account/context)
            mail = outlook.CreateItem(0)  # olMailItem

            # Assign fields
            mail.Subject = subject or ""
            mail.To = "; ".join(to_list)
            mail.CC = "; ".join(cc_list)

            # Assign HTML body safely; fallback to plain Body on failure
            try:
                mail.HTMLBody = body or ""
            except Exception:
                try:
                    # strip HTML tags minimally if needed
                    plain = re.sub(r'<[^>]+>', '', body or "")
                    mail.Body = plain
                except Exception:
                    mail.Body = body or ""

            # Attach files safely
            for att in attachments or []:
                try:
                    if att and os.path.exists(att):
                        mail.Attachments.Add(att)
                except Exception:
                    # skip problematic attachment
                    continue

            # Try to set SaveSentMessageFolder to the default Sent of the account
            # We avoid SendUsingAccount to prevent 4096; rely on default account for sending.
            try:
                # find an account matching sender_email if possible (best-effort)
                target_acc = None
                try:
                    for acc in session.Accounts:
                        try:
                            if acc.SmtpAddress and acc.SmtpAddress.lower() == sender_email:
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

            ui_set(f"Status: Sending via {sender_email or 'default account'}...")
            mail.Send()
            ui_set("Status: Sent successfully!")
            ui_close(3000)

        except Exception as e:
            # Provide concise error message
            msg = str(e)
            ui_set(f"Status: Failed\n{msg}")
            ui_close(6000)
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

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
