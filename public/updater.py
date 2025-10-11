import os
import requests
import shutil
import threading
import subprocess
from pathlib import Path
import tkinter as tk
from tkinter import ttk

# --- CONFIG ---
ROOT_DIR = Path(r"C:\DSQ Enterprise")  # Local base directory
VERCEL_BASE_URL = "https://dsq-beta.vercel.app/"  # <-- change to your actual site
UPDATE_FILE_URL = VERCEL_BASE_URL + "updates.txt"  # Fetch updates.txt online

TEMP_DIR = ROOT_DIR / "temp"
TEMP_DIR.mkdir(parents=True, exist_ok=True)

MAIN_JS_PATH = ROOT_DIR / "main.js"
LOCAL_ELECTRON = ROOT_DIR / "node_modules" / "electron" / "dist" / "electron.exe"


# --- FILE ACTIONS ---
def download_file(vercel_path, local_path):
    """Download file from Vercel public folder and save to destination"""
    file_url = VERCEL_BASE_URL.rstrip("/") + "/" + vercel_path.lstrip("/")
    temp_path = TEMP_DIR / Path(vercel_path).name
    print(f"⬇️ Downloading {file_url}")

    resp = requests.get(file_url, stream=True)
    resp.raise_for_status()
    with open(temp_path, "wb") as f:
        shutil.copyfileobj(resp.raw, f)

    os.makedirs(local_path.parent, exist_ok=True)
    shutil.move(str(temp_path), str(local_path))
    print(f"✅ Saved to {local_path}")


def delete_file(local_path):
    if local_path.exists():
        local_path.unlink()
        print(f"🗑️ Deleted {local_path}")
    else:
        print(f"⚠️ File not found for deletion: {local_path}")


def rename_file(old_path, new_name):
    if old_path.exists():
        new_path = old_path.parent / new_name
        old_path.rename(new_path)
        print(f"✏️ Renamed to {new_path}")
    else:
        print(f"⚠️ File not found for rename: {old_path}")


# --- CLEAN TEMP ---
def clean_temp():
    if TEMP_DIR.exists():
        for item in TEMP_DIR.iterdir():
            try:
                if item.is_file():
                    item.unlink()
                elif item.is_dir():
                    shutil.rmtree(item)
            except Exception as e:
                print(f"⚠️ Failed to remove {item}: {e}")
        print("🧹 Temp folder cleaned.")


# --- PROCESS UPDATE COMMANDS ---
def process_updates(loader):
    try:
        resp = requests.get(UPDATE_FILE_URL, timeout=10)
        resp.raise_for_status()
        lines = [line.strip() for line in resp.text.splitlines() if line.strip()]

        if not lines:
            loader.label.config(text="No updates found.")
            loader.update_progress(100)
            loader.root.after(2000, loader.close)
            return

        total = len(lines)
        done = 0

        for line in lines:
            try:
                if line.lower().startswith("download"):
                    # Example: Download public/app/abc.exe>root/folderA/abc.exe
                    src, dest = line.split(">", 1)
                    vercel_path = src.split(" ", 1)[1].strip()
                    rel_path = dest.replace("root/", "").strip()
                    local_path = ROOT_DIR / rel_path
                    download_file(vercel_path, local_path)

                elif line.lower().startswith("delete"):
                    # Example: Delete root/folderX/xyz.exe
                    rel_path = line.split(" ", 1)[1].replace("root/", "").strip()
                    local_path = ROOT_DIR / rel_path
                    delete_file(local_path)

                elif line.lower().startswith("rename"):
                    # Example: Rename root/folderX/xyz.exe>pqr.exe
                    src, new_name = line.split(">", 1)
                    rel_path = src.split(" ", 1)[1].replace("root/", "").strip()
                    old_path = ROOT_DIR / rel_path
                    rename_file(old_path, new_name.strip())

                else:
                    print(f"⚠️ Unknown command: {line}")

            except Exception as e:
                print(f"❌ Error processing line '{line}': {e}")

            # Update progress bar
            done += 1
            percent = int((done / total) * 100)
            loader.update_progress(percent)

        clean_temp()
        loader.label.config(text="Update complete!\nClose and Restart Application")
        loader.update_progress(100)
        loader.root.after(3000, loader.close)

    except Exception as e:
        loader.label.config(text=f"Error fetching updates: {e}")
        loader.update_progress(100)
        loader.root.after(5000, loader.close)


# --- UI CLASS ---
class UpdaterWindow:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("DSQ Enterprise Updater")
        self.root.geometry("420x140")
        self.root.resizable(False, False)
        self.root.attributes("-topmost", True)
        self.root.eval('tk::PlaceWindow . center')

        self.label = tk.Label(self.root, text="Checking for updates...", font=("Arial", 12))
        self.label.pack(pady=10)

        self.progress = ttk.Progressbar(self.root, mode='determinate', length=350)
        self.progress.pack(pady=10)

        self.percent_label = tk.Label(self.root, text="0%", font=("Arial", 10))
        self.percent_label.pack()

    def update_progress(self, value):
        self.progress["value"] = value
        self.percent_label.config(text=f"{int(value)}%")
        self.root.update_idletasks()

    def close(self):
        self.root.destroy()


# --- MAIN ---
def main():
    loader = UpdaterWindow()
    threading.Thread(target=process_updates, args=(loader,), daemon=True).start()
    loader.root.mainloop()


if __name__ == "__main__":
    main()
