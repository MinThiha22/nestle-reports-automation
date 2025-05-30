import win32com
import win32com.client as win32
import os
import time
import logging
import pythoncom
import threading
import keyboard
from datetime import datetime, timedelta
import ctypes
import sys
import gc

# Import the filtering function from the other file
from NPDFilterProducts import filter_products_by_week_values_v2, excel_date_to_datetime


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
timestamp = datetime.now().strftime('%Y-%m-%d_%H:%M')
error_filename = f"NPD_errors_{timestamp}.log"

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



# =========================================
# Main function to automate Excel process |
# =========================================

file_path = r"C:\Users\chg\OneDrive\NESTLE\NPD.xlsx" # local path to the file
#file_path = r"C:\Users\NZShallaZu\NESTLE\Commercial Development - Documents\General\03 Shopper Centricity\Circana\Circana Pivots\Circana Pivot.xlsx"  # Sharepoint path
def automate_excel_process(file_path):
    today = datetime.now().strftime('%Y-%m-%d')
    print(f"toadayDate: {today}")
    global stop_requested
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


        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")

        print(f"Opening workbook: {file_path}")
        wb = excel.Workbooks.Open(file_path, UpdateLinks=0, ReadOnly=False)
        sheet = wb.Sheets("Actuals") # select sheet Actuals
        pivot = sheet.PivotTables("PivotTable1")
        print("Workbook opened successfully.")
        print("\n\n== 1. Refreshing Pivot Tables ==\n")
        try:
            pivot.PivotCache().Refresh()
            excel.CalculateUntilAsyncQueriesDone()
            print("\nPivot tables refreshed.")
            if stop_requested:
                raise KeyboardInterrupt()
        except Exception as e:
            log_error(f"Pivot table refresh error: {e}")
            print(f"❌ Error refreshing pivot table (logged): {e}")
            
        print("\n\n== 2. Updating Weeks Date ==")
        try:
            # Fields that may have filters or slicers applied
            print("\nClearing filters on slicers and fields...")
            slicer_fields = [
                "[TotalMarket].[Business Unit Value].[Business Unit Value]",
                "[TotalMarket].[Brand Value].[Brand Value]",
                "[TotalMarket].[Geography].[Geography]",
                "[TotalMarket].[Product].[Product]", # we clear to check new products
                "[Table1].[Weeks].[Weeks]", # select all weeks
            ]

            # Clear filters for each slicer-related field
            for field_name in slicer_fields:
                try:
                    field = pivot.PivotFields(field_name)
                    field.ClearAllFilters()
                    print(f"Cleared filters on {field_name}")
                except Exception as e:
                    print(f"Warning: Could not clear filter on {field_name}: {e}")
        
            # Now handle the Weeks field
            print("\n\nSelecting all dates on Weeks field...")
            pivot_field = pivot.PivotFields("[Table1].[Weeks].[Weeks]")
            print("\nAll filters cleared.")
            
            # select all items available
            items = [item.Name for item in pivot_field.PivotItems()]
            print("📅 Dates Founds")
            # print all the items founds (optional) 
            print (f"  - {len(items)} items found in Weeks field:")

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
            print("\n\n Selecting only new SKUs...whiting the last 13 weeks\n")
            
            # select file NPDFilterProducts.py
            # Call the product filtering function
            filter_products_by_week_values_v2(
                sheet_name=sheet,
                pivot_table_name=pivot
            )
            print("✅ Products filtered successfully.")
            
        except Exception as e:
            log_error(f"OLAP pivot interaction failed: {e}")
            print(f"❌ OLAP pivot interaction failed (logged): {e}")

        # Check if the user has requested to stop the process
        if stop_requested:
            raise KeyboardInterrupt()

        print("\n== 3. Refresh Power BI Pivot Table ==")
        try:
            sheet = wb.Sheets("For Power BI") # select sheet For Power BI
            pivot = sheet.PivotTables("PivotTable2")
            pivot.PivotCache().Refresh()
            pivot.RefreshTable()
            print("✅ Power BI Pivot Table refreshed successfully.")

        except Exception as e:
            log_error(f"Power BI Pivots error: {e}")
            print("❌ Power BI Pivots failed (logged)")

        except Exception as e:
            log_error(f"Power BI sheet error: {e}")
            print("❌ Error accessing Power BI pivots (logged)")


        print("\n== Finalizing & Saving ==")
        
        # ===================================================
        # | END AUTOMATE PROCESS, REST SAVE & CLEANUP EXCEL |
        # ===================================================
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
    timer_thread = threading.Thread(target=stopwatch, daemon=True)

    esc_thread.start()
    timer_thread.start()

    automate_excel_process(file_path)
    
    # Ensure logs are flushed before checking file size
    logging.shutdown()

    # Clean up the log file if it's empty
    if os.path.exists(error_filename) and os.path.getsize(error_filename) == 0:
        os.remove(error_filename)
        print("🧹 No errors occurred. Log file deleted.")
    else:
        print(f"⚠️ Errors were logged to: {error_filename}")