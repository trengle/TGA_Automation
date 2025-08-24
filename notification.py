import smtplib
import logging
import os
import getpass
import time
import glob
from datetime import datetime
from dotenv import load_dotenv
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import threading

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
load_dotenv()

program_name = os.path.basename(__file__)
WATCH_FOLDER = '.'  # Current directory

def setup_env():
    current_user = input("Enter your username: ")
    while True:
        email = input("Enter your Gmail address: ").strip()
        if "@gmail.com" not in email:
            print("Invalid email format. Please try again.")
            continue
        else:
            break

    print(f"Ctrl + click -> https://myaccount.google.com/apppasswords to generate an app password. For 'app name' type {program_name}")
    app_pw = getpass.getpass("Paste your app password here (hidden): ").strip()

    smtp_server = "smtp.gmail.com"

    with open(".env", "w") as f:
        f.write(f"TGA_USERNAME={current_user}\n")
        f.write(f"EMAIL_ADDRESS={email}\n")
        f.write(f"APP_PASSWORD={app_pw}\n")
        f.write(f"SMTP_SERVER={smtp_server}\n")
        f.write(f"TO_EMAIL={email}\n")

    print(".env file created successfully!")
    print("Closing setup...")
    time.sleep(2)

def user_switch():
    TGA_USERNAME = os.getenv("TGA_USERNAME")
    if TGA_USERNAME:
        print(f"Hello! Current user: {TGA_USERNAME}")
        switch = input("Would you like to switch users? (Y/N): ").strip().lower()
        if switch == "y":
            setup_env()
            load_dotenv()
    else:
        print("Hello new user! Let's get you set up.")
        setup_env()
        load_dotenv()

def send_email_notification(file_paths):
    EMAIL_ADDRESS = os.getenv("EMAIL_ADDRESS")
    APP_PASSWORD = os.getenv("APP_PASSWORD")
    SMTP_SERVER = os.getenv("SMTP_SERVER")

    subject = "TGA Notification"
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    file_list = "\n".join([os.path.basename(fp) for fp in file_paths])
    body = f"Your sample is ready!\n\nDetected file(s): \n{file_list}\n\nTimestamp: {timestamp}"
    formatted_email = f"Subject: {subject}\n\n{body}"

    try:
        with smtplib.SMTP(SMTP_SERVER, 587) as conn:
            conn.ehlo()
            conn.starttls()
            conn.login(EMAIL_ADDRESS, APP_PASSWORD)
            conn.sendmail(EMAIL_ADDRESS, EMAIL_ADDRESS, formatted_email)
            logging.info(f"Email sent for file(s):\n" + "\n".join(file_paths))
    except Exception as e:
        logging.error(f"Failed to send email: {e}")

csv_queue = []
queue_lock = threading.Lock()

class CSVWatcher(FileSystemEventHandler):
    def on_created(self, event):
        if event.is_directory or not event.src_path.endswith('.csv'):
            return
        logging.info(f"CSV file detected: {event.src_path}")
        with queue_lock:
            csv_queue.append(event.src_path)

def batch_email_sender():
    while True:
        time.sleep(5)  # Wait for batch window
        with queue_lock:
            if csv_queue:
                files_to_send = csv_queue.copy()
                csv_queue.clear()
                send_email_notification(files_to_send)

def start_daemon():
    user_switch()
    threading.Thread(target=batch_email_sender, daemon=True).start()
    event_handler = CSVWatcher()
    observer = Observer()
    observer.schedule(event_handler, WATCH_FOLDER, recursive=False)
    observer.start()
    logging.info(f"Watching for CSV files in: {os.path.abspath(WATCH_FOLDER)}")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
        logging.info("Daemon stopped by user.")
    observer.join()

if __name__ == "__main__":
    start_daemon()
