import win32com.client as win32
import os
import time
import logging
import pythoncom
import threading
import keyboard
from datetime import datetime
import ctypes
import sys

# Prevent sleep (ES_CONTINUOUS | ES_SYSTEM_REQUIRED)
ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001
ES_DISPLAY_REQUIRED = 0x00000002  # Optional: keep screen awake too

def prevent_sleep():
    ctypes.windll.kernel32.SetThreadExecutionState(
        ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED
    )

def allow_sleep():
    ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)

# Setup error logging
timestamp = datetime.now().strftime('%Y%m%d_%H%M')
error_filename = f"circana_errors_{timestamp}.log"

logging.basicConfig(
    filename=error_filename,
    level=logging.ERROR,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

stop_requested = False
start_time = None

def log_error(msg):
    logging.error(msg)

# escape key listener, to exit the script
def listen_for_escape():
    global stop_requested
    print("\n\n======================================")
    print("| Press ESC at any time to cancel... |")
    print("======================================\n\n")
    keyboard.wait('esc')
    stop_requested = True
    print("\n❌ ESC pressed. Cancelling process...")

# timer function, to show elapsed time
def stopwatch():
    global stop_requested, start_time
    start_time = time.time()
    while not stop_requested:
        elapsed = int(time.time() - start_time)
        mins, secs = divmod(elapsed, 60)
        timer = f"⏱️_Timer: {mins:02d}:{secs:02d}"
        print(timer, end='\r')  # Overwrites the same line
        time.sleep(1)
    # Print final time when stopped
    elapsed = int(time.time() - start_time)
    mins, secs = divmod(elapsed, 60)
    print(f"\n⏱️_Total time: {mins} minutes {secs} seconds")


# Main function to automate Excel process
def automate_excel_process():
    global stop_requested, start_time
    try:
        pythoncom.CoInitialize()

        try:
            excel = win32.GetActiveObject('Excel.Application')
            print("Connected to running Excel instance")
        except:
            excel = win32.Dispatch('Excel.Application')
            print("Created new Excel instance")

        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        excel.EnableEvents = False

        file_path = input("Please enter the full file path to the Excel workbook: ").strip().strip('"')
        
        ##file_path = r"C:\Users\chg\OneDrive\NESTLE\Circana Pivot.xlsx" # local path to the file
        # file_path = r"C:\Users\NZShallaZu\NESTLE\Commercial Development - Documents\General\03 Shopper Centricity\Circana\Circana Pivots\Circana Pivot.xlsx"  # Sharepoint path

        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")

        start_time = time.time()
        timer_thread = threading.Thread(target=stopwatch, daemon=True)
        timer_thread.start()
        
        print(f"Opening workbook: {file_path}")
        wb = excel.Workbooks.Open(file_path, UpdateLinks=0, ReadOnly=False)

        print("\n\n== 1. Refreshing all connections ==")
        for conn in wb.Connections:
            if stop_requested:
                raise KeyboardInterrupt()
            try:
                print(f"Refreshing connection: {conn.Name}")
                conn.Refresh()
                time.sleep(2)
            except Exception as e:
                log_error(f"Error refreshing connection '{conn.Name}': {str(e)}")

        if stop_requested:
            raise KeyboardInterrupt()

        print("\n\n== 2. Updating Date Pivot ==")
        try:
            dates_sheet = wb.Sheets("Dates") # Sheet name
            pivot = dates_sheet.PivotTables("PivotTable2") # Pivot table name
            field_name = "[TSM].[Date].[Date]" # Field name to filter (OLAP)
            pivot_field = pivot.PivotFields(field_name) # Get the pivot field 

            # This works for OLAP pivots
            pivot_field.ClearAllFilters() # Clear all filters
            # select all items available
            items = [item.Name for item in pivot_field.PivotItems()]
            print("📅 Dates Founds")
            # print all the items founds (optional) 
            # for name in items:
            #     print(f"  - {name}")

            # filter and Keep all non-blank values
            non_blank_items = [
                d for d in items 
                if d and not d.strip().endswith(".&") and d.strip() != ""
            ]
            if not non_blank_items:
                raise Exception("No non-blank dates found.")
            
            # display the selected (filter items) only non-blank values
            pivot_field.VisibleItemsList = non_blank_items
            
            print(f"✅ Dates updated successfully.")

        except Exception as e:
            log_error(f"OLAP pivot interaction failed: {e}")
            print(f"❌ OLAP pivot interaction failed (logged): {e}")


        # Check if the user has requested to stop the process
        if stop_requested:
            raise KeyboardInterrupt()

        print("\n== 3. Updating Nespresso Pivots ==")
        try:
            # Refresh all Power Queries + OLAP data model + Pivots
            #wb.RefreshAll() alternative to refresh all connections again

            # recalculate and display update values in pivot tables
            print("Refreshing Nespresso pivot tables...") 
            excel.CalculateUntilAsyncQueriesDone()
        except Exception as e:
            log_error(f"Nespresso Pivots error: {e}")
            print("❌ Nespresso Pivots failed (logged)")

        except Exception as e:
            log_error(f"Nespresso sheet error: {e}")
            print("❌ Error accessing Nespresso pivots (logged)")

        print("\n== Finalizing & Saving ==")
        try:
            excel.Calculation = -4105  # xlCalculationAutomatic
            excel.Calculate()
        except Exception as e:
            log_error(f"Calculation mode reset error: {e}")
            print("⚠️ Excel calculation reset failed, trying manual workbook calculation...")
            try:
                wb.Calculate()
            except Exception as ce:
                log_error(f"Workbook manual calculation failed: {ce}")
                print("❌ Manual calculation also failed (logged)")

        wb.Save()
        wb.Close()
        print("✅ Workbook saved and closed.")

    except KeyboardInterrupt:
        print("🚪 Process interrupted by user.")
        log_error("Process interrupted by ESC key.")
        try:
            wb.Close(SaveChanges=False)
        except:
            pass
    except Exception as e:
        log_error(f"Critical error: {e}")
        print(f"❌ Critical error occurred (logged): {e}")
    finally:
        try:
            excel.Application.Quit()
        except:
            pass
        stop_requested = True  # Force timer and ESC thread to stop
        allow_sleep()  # Let the computer sleep again

if __name__ == "__main__":
    prevent_sleep()
    esc_thread = threading.Thread(target=listen_for_escape, daemon=True)
    
    esc_thread.start()

    automate_excel_process()
    
    # Ensure logs are flushed before checking file size
    logging.shutdown()

    # Clean up the log file if it's empty
    if os.path.exists(error_filename) and os.path.getsize(error_filename) == 0:
        os.remove(error_filename)
        print("🧹 No errors occurred. Log file deleted.")
    else:
        print(f"⚠️ Errors were logged to: {error_filename}")