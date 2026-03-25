import time
import pandas as pd
import json
import subprocess
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from datetime import datetime

# === CONFIG ===
EXCEL_FILE = r"C:\Users\jaybe\OneDrive\BFRV AirBNB\2026\2026_airbnb_summary_v1.1.xlsx"
WATCH_FOLDER = r"C:\Users\jaybe\OneDrive\BFRV AirBNB\2026"

LAST_RUN = 0
DEBOUNCE_SECONDS = 5


def wait_until_file_ready(filepath):
    """Wait until Excel releases the file (no spam, clean wait)"""
    while True:
        try:
            with open(filepath, 'rb'):
                return
        except PermissionError:
            print("⏳ Waiting for Excel to finish...")
            time.sleep(2)


def process_file():
    global LAST_RUN

    now = time.time()
    if now - LAST_RUN < DEBOUNCE_SECONDS:
        return

    LAST_RUN = now

    print(f"\n[{datetime.now()}] Change detected")

    try:
        # Wait until file is no longer locked
        wait_until_file_ready(EXCEL_FILE)

        # Read Excel
        df = pd.read_excel(EXCEL_FILE)
        data = df.to_dict(orient="records")

        # Save JSON
        with open("data.json", "w") as f:
            json.dump(data, f, indent=2)

        print("✅ JSON updated")

        # Git commands
        subprocess.run("git add data.json", shell=True)
        subprocess.run('git diff --quiet || git commit -m "Auto update"', shell=True)
        subprocess.run("git push", shell=True)

        print("🚀 Pushed to GitHub")

    except Exception as e:
        print("❌ ERROR:", e)


class ExcelHandler(FileSystemEventHandler):
    def on_modified(self, event):
        if event.src_path.endswith(".xlsx"):
            process_file()


if __name__ == "__main__":
    observer = Observer()
    observer.schedule(ExcelHandler(), path=WATCH_FOLDER, recursive=False)
    observer.start()

    print("👀 Watching Excel folder... (Press Ctrl+C to stop)")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()

    observer.join()