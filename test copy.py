import win32com.client
from datetime import datetime, timedelta, date
import openpyxl
import openpyxl.utils

def excel_date_to_datetime(excel_date):
    if excel_date is None or excel_date == "":
        return None
    try:
        return datetime(1899, 12, 30) + timedelta(days=excel_date)
    except TypeError:
        return None

def filter_products_by_week_values_v2(filename, sheet_name, pivot_table_name, product_column_letter='E', first_product_row=21, weeks_header_row=20, first_week_column_letter='G'):
    excel = None
    workbook = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True
        workbook = excel.Workbooks.Open(filename)
        worksheet = workbook.Sheets(sheet_name)
        pivot = worksheet.PivotTables(pivot_table_name)

        if not pivot:
            print(f"Pivot Table '{pivot_table_name}' not found.")
            return

        product_field = None
        for pf in pivot.PivotFields():
            if pf.Caption == "Product":
                product_field = pf
                break

        if not product_field:
            print("PivotField 'Product' not found.")
            return

        # 1. Find the last "Weeks" column (by finding the last non-empty cell in the header row)
        # Calculate the date 13 weeks ago
        # today = date.today()
        today = '20/04/2025'  # For testing purposes, set a fixed date
        # Convert string to date
        today = datetime.strptime(today, '%d/%m/%Y').date() #for testing
        sunday = today - timedelta(days=(today.weekday())+1) # this will give us the last Sunday, +1 to go back to the previous Sunday
        date_13_weeks_ago = sunday - timedelta(weeks=13)
        cutoff_date_str = date_13_weeks_ago.strftime('%d/%m/%Y') # Format to match Excel

        cutoff_col_index = -1
        last_week_col_index = -1
        header_row = weeks_header_row
        start_week_col_index = openpyxl.utils.column_index_from_string(first_week_column_letter)
        
        # Search for the last week column
        for col_index in range(start_week_col_index, 500): # Increased upper bound
            if worksheet.Cells(header_row, col_index).Value is not None:
                last_week_col_index = col_index
            else:
                break

        # Search for the column with the cutoff date
        for col_index in range(start_week_col_index, 200):
            cell_value = worksheet.Cells(header_row, col_index).Value
            if cell_value is not None:
                try:
                    cell_date = None
                    if isinstance(cell_value, (int, float)):
                        cell_date = excel_date_to_datetime(cell_value)
                    elif isinstance(cell_value, datetime):
                        cell_date = cell_value

                    if cell_date:
                        if cell_date.strftime('%d/%m/%Y') == cutoff_date_str:
                            cutoff_col_index = col_index
                            break
                except:
                    pass
            else:
                break

        if cutoff_col_index == -1:
            print(f"Warning: Could not find the column for the date '{cutoff_date_str}'. Using a fallback.")
            # Fallback to 13 columns before the last week found earlier
            cutoff_col_index = max(start_week_col_index, last_week_col_index - 13)

        print(f"13-Week Cutoff Column (by date): {openpyxl.utils.get_column_letter(cutoff_col_index)}")

        # 3. Iterate through products and check for values
        product_col_index = openpyxl.utils.column_index_from_string(product_column_letter)
        last_product_row = pivot.TableRange1.Row + pivot.TableRange1.Rows.Count
        products_to_select = []
        previous_product = None

        # Move the cutoff column one week earlier
        adjusted_cutoff_col_index = max(start_week_col_index, cutoff_col_index - 1)
        print(f"Adjusted Cutoff Column: {openpyxl.utils.get_column_letter(adjusted_cutoff_col_index)} (Index: {adjusted_cutoff_col_index})")

        print(f"cut off col index: {cutoff_col_index}, last week col index: {last_week_col_index}")
        for row in range(first_product_row, last_product_row + 1):
            product_name = worksheet.Cells(row, product_col_index).Value
            if product_name and product_name != previous_product:
                has_value_before_cutoff = False
                for col in range(start_week_col_index, adjusted_cutoff_col_index + 1):
                    value = worksheet.Cells(row, col).Value
                    if value is not None and value != "":
                        has_value_before_cutoff = True
                        break

                if not has_value_before_cutoff:
                    # Check for at least 2 values in the 13 weeks after the (original) cutoff
                    sales_count_after_cutoff = 0
                    for col_after in range(cutoff_col_index, last_week_col_index + 1):
                        value_after = worksheet.Cells(row, col_after).Value
                        print(f"{value_after} values")
                        if value_after is not None:
                            print(f"in If Loop {value_after} values")
                            sales_count_after_cutoff += 1
                            print(f"PRINT has sales after the original cutoff: {sales_count_after_cutoff}.")

                    if sales_count_after_cutoff >= 2:
                        products_to_select.append(f"[TotalMarket].[Product].&[{product_name}]")
                        # print(f"Selecting product '{product_name}' (no sales before adjusted cutoff and 2+ after original).")
                    # else:
                        # print(f"Product '{product_name}' has < 2 sales after the original cutoff. Not selecting.")
                # else:
                    # print(f"Product '{product_name}' has sales before the adjusted cutoff. Not selecting.")

                previous_product = product_name
            elif product_name == previous_product:
                continue

        # 4. Apply the filter
        if products_to_select:
            product_field.VisibleItemsList = products_to_select
            print(f"Selected {len(products_to_select)} products with no history before the cutoff.")
        else:
            print("No products found that match the criteria.")

        workbook.Save()
        workbook.Close()
        excel.Quit()
        excel = None
        workbook = None

    except Exception as e:
        print(f"An error occurred: {e}")
        if workbook:
            workbook.Close(SaveChanges=False)
        if excel:
            excel.Quit()
            excel = None
        workbook = None

if __name__ == '__main__':
    filename = r"C:\Users\chg\OneDrive\NESTLE\NPD.xlsx"
    sheet_name = "Actuals"
    pivot_table_name = "PivotTable1"
    product_column = 'E'
    first_product_row_num = 21
    weeks_header_row_num = 20
    first_week_column = 'G'

    filter_products_by_week_values_v2(filename, sheet_name, pivot_table_name, product_column, first_product_row_num, weeks_header_row_num, first_week_column)