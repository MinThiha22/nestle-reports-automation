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

def filter_products_by_week_values_v2(sheet_name, pivot_table_name, product_column_letter='E',first_product_row=21,weeks_header_row=20,first_week_column_letter='G'):
    worksheet = sheet_name
    pivot = pivot_table_name
   

    try:
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
    
        print("PivotField 'Product' found. Proceeding with filtering...")
        # 1. Find the last "Weeks" column (by finding the last non-empty cell in the header row)
        # Calculate the date 13 weeks ago
        # today = '20/04/2025'  # For testing purposes, set a fixed date
        # today = datetime.strptime(today, '%d/%m/%Y').date() #for testing
        today = date.today()
        sunday = today - timedelta(days=(today.weekday())+1) # this will give us the last Sunday, +1 to go back to the previous Sunday
        date_13_weeks_ago = sunday - timedelta(weeks=14) # system is behinf 1 week, so we need to go back 14 weeks to get the correct date
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

        # 3. Iterate through products and check for values before the cutoff
        product_col_index = openpyxl.utils.column_index_from_string(product_column_letter)
        last_product_row = pivot.TableRange1.Row + pivot.TableRange1.Rows.Count
        products_to_select = set() # Use a set for efficient checking and to avoid duplicates
        banned_items = set()
        
        # Move the cutoff column one week earlier
        adjusted_cutoff_col_index = max(start_week_col_index, cutoff_col_index - 1)
        print(f"Adjusted Cutoff Column: {openpyxl.utils.get_column_letter(adjusted_cutoff_col_index)} (Index: {adjusted_cutoff_col_index})")

        for row in range(first_product_row, last_product_row + 1):
            product_name = worksheet.Cells(row, product_col_index).Value
            if product_name and product_name not in banned_items:
                has_value_before_cutoff = False
                for col in range(start_week_col_index, adjusted_cutoff_col_index + 1):
                    value = worksheet.Cells(row, col).Value
                    if value is not None and value != "":
                        has_value_before_cutoff = True
                        break

                sales_count_after_cutoff = 0
                for col_after in range(cutoff_col_index, last_week_col_index + 1):
                    value_after = worksheet.Cells(row, col_after).Value
                    if value_after is not None and value_after > 0:
                        sales_count_after_cutoff += 1

                meets_criteria = not has_value_before_cutoff and sales_count_after_cutoff >= 1

                if meets_criteria:
                    products_to_select.add(f"[TotalMarket].[Product].&[{product_name}]")
                    print(f"Product '{product_name}' meets criteria. Added to selection.")
                else:
                    if f"[TotalMarket].[Product].&[{product_name}]" in products_to_select:
                        products_to_select.discard(f"[TotalMarket].[Product].&[{product_name}]")
                        # print(f"Product '{product_name}' no longer meets criteria. Removed from selection.")
                    banned_items.add(product_name)
                    # print(f"Product '{product_name}' did not meet criteria. Added to banned list.")

        final_products_to_select = list(products_to_select) # Convert back to list for setting filter


        # 4. Apply the filter
        if final_products_to_select:
            product_field.VisibleItemsList = final_products_to_select
            print(f"Selected {len(final_products_to_select)} products.")
        else:
            print("No products found that meet the criteria across all occurrences.")

    except Exception as e:
        print(f"An error occurred: {e}")
