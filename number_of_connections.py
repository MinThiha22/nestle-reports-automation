import win32com.client
import os

file_path = r"C:\Users\chg\OneDrive\NESTLE\Circana Pivot.xlsx"  # Update to your actual path

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(os.path.abspath(file_path))
    print(f"Opened: {file_path}")

    for conn in wb.Connections:
        print(f"\nName: {conn.Name}")
        print(f"Type: {conn.Type}")  # 1 = OLEDB, 2 = ODBC, 7 = Power Query

        try:
            print(f"  BackgroundQuery: {conn.OLEDBConnection.BackgroundQuery}")
        except:
            print("  BackgroundQuery: N/A")

except Exception as e:
    print(f"Error: {e}")
finally:
    wb.Close(SaveChanges=False)
    excel.Quit()
