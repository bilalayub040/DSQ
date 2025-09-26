import os
import requests
import subprocess
import tkinter as tk

# URLs on Vercel
BASE_URL = "https://dsq-beta.vercel.app"
EXE_URL = f"{BASE_URL}/launcher.exe"
VERSION_URL = f"{BASE_URL}/launcher.version"

# Local paths
LOCAL_EXE = os.path.join(os.getcwd(), "launcher.exe")
LOCAL_VERSION = os.path.join(os.getcwd(), "launcher.version")

# GUI setup
root = tk.Tk()
root.overrideredirect(True)  # no title bar
root.geometry("300x100+600+350")
root.configure(bg="#f0f0f0")

# Labels
start_label = tk.Label(root, text="Starting the App.", font=("Arial", 12), bg="#f0f0f0")
start_label.pack(pady=(20, 5))

check_label = tk.Label(root, text="Checking updates", font=("Arial", 10), bg="#f0f0f0")
check_label.pack()

# For animated dots
dots = 0
def animate_dots():
    global dots
    dots = (dots + 1) % 4
    check_label.config(text="Checking updates" + "." * dots)
    root.update_idletasks()
    root.after(500, animate_dots)

def get_remote_version():
    try:
        response = requests.get(VERSION_URL, timeout=10)
        if response.status_code == 200:
            return response.text.strip()
    except:
        pass
    return None

def get_local_version():
    if os.path.exists(LOCAL_VERSION):
        with open(LOCAL_VERSION, "r") as f:
            return f.read().strip()
    return None

def download_file(url, local_path):
    try:
        response = requests.get(url, stream=True, timeout=30)
        response.raise_for_status()
        with open(local_path, "wb") as f:
            for chunk in response.iter_content(chunk_size=8192):
                f.write(chunk)
        return True
    except:
        return False

def update_and_run():
    remote_version = get_remote_version()
    local_version = get_local_version()

    if remote_version and remote_version != local_version:
        download_file(EXE_URL, LOCAL_EXE)
        with open(LOCAL_VERSION, "w") as f:
            f.write(remote_version)

    # Launch the app
    if os.path.exists(LOCAL_EXE):
        subprocess.Popen([LOCAL_EXE], shell=True)

    # Close the updater window immediately
    root.destroy()

# Start the animated dots
animate_dots()

# Start update process after short delay to let animation start
root.after(100, update_and_run)

root.mainloop()
