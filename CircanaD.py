import os
import shutil
from selenium import webdriver
from selenium.webdriver.edge.service import Service
import time
import logging
import threading
import ctypes
from datetime import datetime
import keyboard
import win32com.client as win32
import pythoncom
import sys
import winshell

# ========== Configuration ========== #
download_path = r"C:\Users\chg\Desktop\Download_01"
destination_folder = r"C:\Users\chg\Desktop\Circana_Flat_Files"
# download_path = r"C:\Users\NZShallaZu\Downloads"
# destination_folder = r"C:\Users\NZShallaZu\NESTLE\Commercial Development - Documents\General\03 Shopper Centricity\Circana\Flat Files"

ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001
ES_DISPLAY_REQUIRED = 0x00000002

stop_requested = False
start_time = None

# ========== Utilities ========== #
def prevent_sleep():
    ctypes.windll.kernel32.SetThreadExecutionState(
        ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED
    )

def allow_sleep():
    ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)

def log_error(msg):
    logging.error(msg)

def listen_for_escape():
    global stop_requested
    print("\n\n======================================")
    print("| Press ESC at any time to cancel... |")
    print("======================================\n\n")
    keyboard.wait('esc')
    stop_requested = True
    print("\n❌  ESC pressed. Cancelling process...")

def stopwatch():
    global stop_requested, start_time
    start_time = time.time()
    while not stop_requested:
        elapsed = int(time.time() - start_time)
        mins, secs = divmod(elapsed, 60)
        print(f"⏱️  Timer: {mins:02d}:{secs:02d}", end='\r')
        time.sleep(1)
    elapsed = int(time.time() - start_time)
    mins, secs = divmod(elapsed, 60)
    print(f"\n⏱️  Total time: {mins} minutes {secs} seconds")

# ========== Main Automation Logic ========== #
def run_excel_updates():
    pythoncom.CoInitialize()

    try:
        try:
            excel = win32.GetActiveObject("Excel.Application")
            print("Connected to running Excel instance")
        except Exception:
            excel = win32.Dispatch("Excel.Application")
            print("Created new Excel instance")

        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        excel.EnableEvents = False

        # === Calendar Sheet Update ===
        print("\n\n=== Updating Calendar Sheet ===")
        calendar_file_path = os.path.join(destination_folder, "Calendar.xlsx")
        if not os.path.exists(calendar_file_path):
            raise FileNotFoundError(f"❌  File not found: {calendar_file_path}")
        wb = excel.Workbooks.Open(calendar_file_path, UpdateLinks=0, ReadOnly=False)
        sheet = wb.Sheets("Sheet1")
        date_value = sheet.Range("J5").Value
        sheet.Range("J1").Value = date_value
        print(f"📅  Calendar.xlsx: J1 updated to {date_value}")
        wb.Save()
        wb.Close(SaveChanges=True)

        # === Week Sheet Update ===
        print("\n\n=== Updating Week Sheet ===")
        week_file_path = os.path.join(destination_folder, "Weeks.xlsx")
        if not os.path.exists(week_file_path):
            raise FileNotFoundError(f"❌  File not found: {week_file_path}")
        wb = excel.Workbooks.Open(week_file_path, UpdateLinks=0, ReadOnly=False)
        sheet = wb.Sheets("Sheet1")
        date_value = sheet.Range("P2").Value
        sheet.Range("A2").Value = date_value
        print(f"📅  Weeks.xlsx: A2 updated to {date_value}")
        wb.Save()
        wb.Close(SaveChanges=True)

    finally:
        excel.Quit()
        
# ========== Clean Recycle Bin ========== #        
recycle_bin_cleaned = False

def clean_recycle_bin():
    global recycle_bin_cleaned
    try:
        pythoncom.CoInitialize()
        if not recycle_bin_cleaned:
            winshell.recycle_bin().empty(confirm=False, show_progress=False, sound=False)
            recycle_bin_cleaned = True
            print("🗑️  Recycle Bin cleaned.")
    except Exception as e:
        print(f"❌  Failed to clean Recycle Bin: {e}")
        log_error(f"❌  Failed to clean Recycle Bin: {e}")
    finally:
        pythoncom.CoUninitialize()


# ========== Main Function ========== #
def main():
    global stop_requested
    timestamp = datetime.now().strftime('%Y%m%d_%H%M')
    error_filename = f"circana_errors_{timestamp}.log"

    logging.basicConfig(
        filename=error_filename,
        level=logging.ERROR,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )

    prevent_sleep()
    esc_thread = threading.Thread(target=listen_for_escape, daemon=True)
    timer_thread = threading.Thread(target=stopwatch, daemon=True)
    esc_thread.start()
    timer_thread.start()

    print("🚀  Starting Circana Automation...")

    try:
        # Step 1: Clean download folder
        print("\n\n=== Deleting files in download folder ===")
        for filename in os.listdir(download_path):
            if stop_requested: raise KeyboardInterrupt()
            file_path = os.path.join(download_path, filename)
            if os.path.isfile(file_path):
                os.remove(file_path)
                print(f"Deleted: {file_path}")

        # Step 2: Download from website
        # Skipped...

        # Step 3: Clean destination folder of old duplicates
        print("\n\n=== Cleaning duplicates in destination folder ===")
        try:
            if os.path.exists(destination_folder):
                download_filenames = set(os.listdir(download_path))  # freshly downloaded files
                for filename in download_filenames:
                    dest_file_path = os.path.join(destination_folder, filename)
                    if stop_requested:
                        raise KeyboardInterrupt()
                    if os.path.isfile(dest_file_path):
                        os.remove(dest_file_path)
                        print(f"🗑️ Deleted old version from destination: {dest_file_path}")
        except Exception as e:
            print(f"❌  Error cleaning duplicates in destination folder: {e}")
            log_error(f"❌  Error cleaning duplicates: {e}")
                     
        # step 4: remove files from Recycle Bin
        print("\n\n=== Cleaning Recycle Bin ===")
        clean_recycle_bin()

        # Step 5: Move files to destination
        print("\n\n=== Moving files to destination folder ===")
        for filename in os.listdir(download_path):
            if stop_requested: raise KeyboardInterrupt()
            src = os.path.join(download_path, filename)
            dst = os.path.join(destination_folder, filename)
            shutil.move(src, dst)
            print(f"Moved: {filename} ➜ {destination_folder}")

        # Step 6 & 7: Excel updates
        run_excel_updates()

    except KeyboardInterrupt:
        print("🛑  Process manually interrupted.")
    except Exception as e:
        log_error(f"Unhandled error: {e}")
        print(f"❌  Error: {e}")
    finally:
        allow_sleep()
        logging.shutdown()

        if os.path.exists(error_filename) and os.path.getsize(error_filename) == 0:
            os.remove(error_filename)
            print("🧹  No errors occurred. Log file deleted.")
        else:
            print(f"⚠️  Errors were logged to: {error_filename}")

        print("✅  Circana Automation Complete.")

if __name__ == "__main__":
    main()