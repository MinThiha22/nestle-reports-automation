import gc
import os
import sys
import time
import platform
import ctypes
import pythoncom
import win32com.client as win32
import win32process
import io

# === UTF-8 output for Electron logs ===
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001
ES_DISPLAY_REQUIRED = 0x00000002 # Optional: keep screen awake too

excel_pid = None

# === Helper functions ===
def force_console_output(message):
    """Print and flush immediately (so Electron sees it)."""
    try:
        print(message)
    except UnicodeEncodeError:
        print(message.encode("utf-8", errors="replace").decode("utf-8"))
    sys.stdout.flush()
    sys.stderr.flush()

def prevent_sleep():
    if platform.system() == "Windows":
        ctypes.windll.kernel32.SetThreadExecutionState(
            ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED
        )

def allow_sleep():
    if platform.system() == "Windows":
        ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)

def start_excel():
    """Start Excel and track PID for external kill."""
    global excel_pid
    excel_app = win32.DispatchEx("Excel.Application")
    excel_app.Visible = True  # 👈 make the window visible
    hwnd = excel_app.Hwnd
    _, excel_pid = win32process.GetWindowThreadProcessId(hwnd)
    return excel_app

# === Main automation function ===
def automate(source, target):
  global excel_pid
  
  try:
    excel = start_excel()
    excel.Visible = True
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    excel.EnableEvents = False
    
    if not os.path.exists(source):
      raise FileNotFoundError(f"[Source] File not found: {source}")
    
    if not os.path.exists(target):
      raise FileNotFoundError(f"[Target] File not found: {target}")
    
    # Open workbooks
    force_console_output(f"[Source] 🔄 Opening workbook: {os.path.basename(source)}")
    wb_src = excel.Workbooks.Open(source, UpdateLinks=0, ReadOnly=False)
    force_console_output("[Source] ✅ Workbook opened successfully")
    
    force_console_output(f"[Target] 🔄 Opening workbook: {os.path.basename(target)}")
    wb_tar = excel.Workbooks.Open(target, UpdateLinks=0, ReadOnly=False)
    force_console_output("[Target] ✅ Workbook opened successfully")
    
    # Clear target file (Quantium Data)
    ws_tar = wb_tar.Sheets("Channel Level Product Data")
    last_row_tar = ws_tar.UsedRange.Rows.Count + ws_tar.UsedRange.Row - 1
    last_col_tar = ws_tar.UsedRange.Columns.Count + ws_tar.UsedRange.Column - 1
    ws_tar.Range(ws_tar.Cells(3, 3), ws_tar.Cells(last_row_tar, last_col_tar)).ClearContents()
    force_console_output("[Target] ✅ Data cleared successfully")
    
    # Copy data from source and paste on target file
    ws_src = wb_src.Sheets("Channel Level Product Data")
    last_row_src = ws_src.UsedRange.Rows.Count + ws_src.UsedRange.Row - 1
    copy_data = ws_src.Range(f"B8:Q{last_row_src}")
  
    target_start = ws_tar.Cells(2,3) # C2
    copy_data.Copy(target_start)
    force_console_output("✅ Data Copied and Pasted to Target Files")
    
    # Check NA or Zero in formula columns
    error_cells = []
    last_row = ws_tar.UsedRange.Rows.Count + ws_tar.UsedRange.Row - 1
    force_console_output(f"{last_row}")
    values = ws_tar.Range(ws_tar.Cells(2, 1), ws_tar.Cells(last_row, 2)).Value  
    xlErrNA = -2146826246  
    for i, (a, b) in enumerate(values, start=2): 
      if isinstance(a, Exception) or a == xlErrNA:
        error_cells.append(("A", i))

      if isinstance(b, Exception) or b == xlErrNA:
        error_cells.append(("B", i))
    
    if error_cells:
      for col_letter, row in error_cells:
        force_console_output(f"⚠️ N/A found at: {col_letter},{row}")
      force_console_output(f"❌ Total Error cells: {len(error_cells)}")
    else:
      force_console_output("✅ No N/A found in columns A or B")
      
    wb_src.Save()
    wb_src.Close()
    del wb_src
    gc.collect()
    force_console_output("[Source] ✅ Workbook saved and closed")
    
    wb_tar.Save()
    wb_tar.Close()
    del wb_tar
    gc.collect()
    force_console_output("[Target] ✅ Workbook saved and closed")
    excel.Quit()   
    
    return True
      
  except Exception as e:
    force_console_output(f"❌ Critical error: {e}")
    return False

  finally:
    pythoncom.CoUninitialize()
    force_console_output("✅ COM uninitialized")


if __name__ == "__main__":
  
  monthly_report = r"H:\R&D\data\Nestle NZ Online Reporting 20250831.xlsx" # Source
  quantium_data = r"H:\R&D\data\QuantiumData.xlsx" # Target
  automate(monthly_report, quantium_data)
  time.sleep(10)
  