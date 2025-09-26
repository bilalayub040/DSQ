import os
import requests
from pathlib import Path
from urllib.parse import urljoin
import threading
import shutil
import tkinter as tk
from tkinter import ttk

# --- CONFIG ---
BASE_DIR = Path(r"C:\DSQ Enterprise")
TEMP_DIR = BASE_DIR / "temp_updater"
FILE_LIST_URL = "https://dsq-beta.vercel.app/updates.txt"
FILES_BASE_URL = "https://dsq-beta.vercel.app/"  # files hosted at root


class UpdaterWindow:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Updating")
        self.root.geometry("400x120")
        self.root.resizable(False, False)
        self.root.attributes("-topmost", True)
        self.root.eval('tk::PlaceWindow . center')

        self.label = tk.Label(self.root, text="Updating files...", font=("Arial", 12))
        self.label.pack(pady=10)

        self.progress = ttk.Progressbar(self.root, mode='determinate', length=350)
        self.progress.pack(pady=10)

        self.percent_label = tk.Label(self.root, text="0%", font=("Arial", 10))
        self.percent_label.pack()

    def update_progress(self, value):
        self.progress['value'] = value
        self.percent_label.config(text=f"{int(value)}%")
        self.root.update_idletasks()

    def close(self):
        self.progress.stop()
        self.root.destroy()


def download_to_temp(file_name):
    TEMP_DIR.mkdir(parents=True, exist_ok=True)
    temp_file_path = TEMP_DIR / file_name
    file_url = urljoin(FILES_BASE_URL, file_name)
    resp = requests.get(file_url, stream=True, timeout=15)
    resp.raise_for_status()
    with open(temp_file_path, "wb") as f:
        for chunk in resp.iter_content(chunk_size=8192):
            f.write(chunk)
    return temp_file_path


def find_local_files(file_name):
    matches = []
    for root, dirs, files in os.walk(BASE_DIR):
        if file_name in files:
            matches.append(Path(root) / file_name)
    return matches


def background_task(loader):
    try:
        # --- Fetch file list ---
        resp = requests.get(FILE_LIST_URL, timeout=10)
        resp.raise_for_status()
        content = resp.text.strip()
        entries = [f.strip() for f in content.split(",") if f.strip()]

        if not entries:
            loader.label.config(text="No files to update.")
            loader.update_progress(100)
            loader.root.after(1000, loader.close)
            return

        # --- Prepare all file operations ---
        total_operations = 0
        operations = []

        for entry in entries:
            if ">" in entry:
                src_name, dest_name = map(str.strip, entry.split(">"))
            else:
                src_name = dest_name = entry

            local_files = find_local_files(dest_name)
            if local_files:
                for lf in local_files:
                    operations.append((src_name, dest_name, lf))
                    total_operations += 1
            else:
                # File not found locally, we can just place it at BASE_DIR
                operations.append((src_name, dest_name, BASE_DIR / dest_name))
                total_operations += 1

        # --- Execute downloads and replacements ---
        progress_count = 0
        for src_name, dest_name, local_path in operations:
            temp_file = download_to_temp(src_name)

            # Remove original src_name if different from dest_name
            if src_name != dest_name:
                old_src_files = find_local_files(src_name)
                for f in old_src_files:
                    try:
                        f.unlink()
                    except Exception:
                        pass

            # Move temp file to final location (overwrite)
            local_path.parent.mkdir(parents=True, exist_ok=True)
            shutil.move(str(temp_file), local_path)

            progress_count += 1
            percent = (progress_count / total_operations) * 100
            loader.update_progress(percent)

        # --- Cleanup temp folder ---
        if TEMP_DIR.exists():
            shutil.rmtree(TEMP_DIR)

        loader.label.config(text="Update complete!")
        loader.update_progress(100)
        loader.root.after(500, loader.close)

    except Exception as e:
        loader.label.config(text=f"Error: {e}")
        loader.update_progress(100)
        loader.root.after(5000, loader.close)


def main():
    loader = UpdaterWindow()
    threading.Thread(target=background_task, args=(loader,), daemon=True).start()
    loader.root.mainloop()


if __name__ == "__main__":
    main()
