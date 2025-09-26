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


def run_main(loader):
    if not MAIN_JS_PATH.exists():
        print(f"❌ main.js not found at {MAIN_JS_PATH}")
        loader.root.after(0, loader.close)
        return
    if LOCAL_ELECTRON.exists():
        print("Launching main.js with local Electron...")
        subprocess.Popen([str(LOCAL_ELECTRON), str(MAIN_JS_PATH)], cwd=BASE_DIR)
        # Close loader immediately after launch
        loader.root.after(0, loader.close)
    else:
        print("❌ Electron not found in node_modules.")
        loader.root.after(0, loader.close)


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

        self.progress = ttk.Progressbar(self.root, mode='indeterminate', length=300)
        self.progress.pack(pady=10)
        self.progress.start(10)  # fast spinning bar

    def close(self):
        self.progress.stop()
        self.root.destroy()


def main():
    loader = LoadingWindow()
    threading.Thread(target=run_main, args=(loader,), daemon=True).start()
    loader.root.mainloop()


if __name__ == "__main__":
    main()
