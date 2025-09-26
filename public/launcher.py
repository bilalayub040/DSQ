import os
import subprocess
from pathlib import Path
import threading
import tkinter as tk
from tkinter import ttk

# --- CONFIG ---
BASE_DIR = Path(r"C:\DSQ Enterprise")
MAIN_JS_PATH = BASE_DIR / "main.js"
LOCAL_ELECTRON = BASE_DIR / "node_modules" / "electron" / "dist" / "electron.exe"


def run_main():
    if not MAIN_JS_PATH.exists():
        print(f"❌ main.js not found at {MAIN_JS_PATH}")
        return
    if LOCAL_ELECTRON.exists():
        print("Launching main.js with local Electron...")
        # Non-blocking so launcher can close immediately
        subprocess.Popen([str(LOCAL_ELECTRON), str(MAIN_JS_PATH)], cwd=BASE_DIR)
    else:
        print("❌ Electron not found in node_modules. Please ensure it is bundled correctly.")


class LoadingWindow:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("DSQ Enterprise")
        self.root.geometry("400x100")
        self.root.resizable(False, False)
        self.root.attributes("-topmost", True)
        self.root.eval('tk::PlaceWindow . center')

        self.label = tk.Label(self.root, text="Loading the Application...", font=("Arial", 12))
        self.label.pack(pady=10)

        self.progress = ttk.Progressbar(self.root, mode='determinate', length=300)
        self.progress.pack(pady=10)
        self.progress['value'] = 0

    def update_progress(self, percent):
        self.progress['value'] = percent
        self.root.update_idletasks()

    def close(self):
        self.root.destroy()


def background_task(loader):
    # Animate progress from 0 → 90 while preparing
    for i in range(0, 91, 5):
        loader.update_progress(i)
        loader.root.after(50)  # small delay for smooth animation

    # --- Run the app (90-100%) ---
    run_main()
    loader.update_progress(100)

    # Close loader immediately
    loader.root.after(100, loader.close)


def main():
    loader = LoadingWindow()
    threading.Thread(target=background_task, args=(loader,), daemon=True).start()
    loader.root.mainloop()


if __name__ == "__main__":
    main()
