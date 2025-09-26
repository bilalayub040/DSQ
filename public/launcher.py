import os
import sys
import requests
import zipfile
import subprocess
from pathlib import Path
import shutil

# --- CONFIG ---
BASE_DIR = Path(r"C:\DSQ Enterprise")
MAIN_JS_PATH = BASE_DIR / "main.js"
DEPS_ZIP_PATH = BASE_DIR / "deps.zip"
LOCAL_ELECTRON = BASE_DIR / "node_modules" / "electron" / "dist" / "electron.exe"

MAIN_URL = "https://dsq-beta.vercel.app/main.js"
DEPS_URL = "https://dsq-beta.vercel.app/deps.zip"


def is_frozen():
    """Detect if running as frozen exe"""
    return getattr(sys, 'frozen', False)


def download_file(url, dest):
    """Download file from URL to destination"""
    print(f"Downloading {url} -> {dest}")
    resp = requests.get(url, stream=True, timeout=15)
    resp.raise_for_status()
    with open(dest, "wb") as f:
        for chunk in resp.iter_content(chunk_size=8192):
            f.write(chunk)


def needs_update(local_file, url):
    """Check if remote file differs from local (by size)"""
    try:
        head = requests.head(url, timeout=5)
        remote_len = int(head.headers.get("Content-Length", 0))
        if local_file.exists() and remote_len == local_file.stat().st_size:
            return False
        return True
    except Exception as e:
        print(f"[WARN] Could not check update for {url}: {e}")
        return False


def extract_zip(zip_path, extract_to):
    """Extract ZIP directly into the target folder"""
    with zipfile.ZipFile(zip_path, "r") as zf:
        for member in zf.namelist():
            target_path = extract_to / member
            if target_path.exists():
                if target_path.is_file():
                    target_path.unlink()
                else:
                    shutil.rmtree(target_path)
        zf.extractall(extract_to)


def run_main():
    """Run main.js using local Electron"""
    if not MAIN_JS_PATH.exists():
        print(f"❌ main.js not found at {MAIN_JS_PATH}")
        return

    if LOCAL_ELECTRON.exists():
        print("Launching main.js with local Electron...")
        subprocess.run([str(LOCAL_ELECTRON), str(MAIN_JS_PATH)], cwd=BASE_DIR)
    else:
        print("❌ Electron not found in node_modules. Please ensure it is bundled correctly.")


def main():
    BASE_DIR.mkdir(parents=True, exist_ok=True)

    # --- Update main.js ---
    if needs_update(MAIN_JS_PATH, MAIN_URL):
        print("Updating main.js...")
        download_file(MAIN_URL, MAIN_JS_PATH)
    else:
        print("main.js up to date.")

    # --- Update deps.zip ---
    if needs_update(DEPS_ZIP_PATH, DEPS_URL):
        print("Updating deps.zip...")
        download_file(DEPS_URL, DEPS_ZIP_PATH)
        extract_zip(DEPS_ZIP_PATH, BASE_DIR)
    else:
        print("deps.zip up to date.")

    # --- Run the app ---
    run_main()


if __name__ == "__main__":
    main()
