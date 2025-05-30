import win32com.client

def select_single_product(filename, sheet_name, pivot_table_name, product_field_caption="Product", product_to_select="[TotalMarket].[Product].&[NESTLE 100 GRAND BAR 42G]"):
    excel = None
    workbook = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True  # Make Excel visible for debugging
        workbook = excel.Workbooks.Open(filename)
        worksheet = workbook.Sheets(sheet_name)
        pivot = worksheet.PivotTables(pivot_table_name)

        if pivot:
            pivot_field = pivot.PivotFields("[TotalMarket].[Product].[Product]")
            print("\nAll filters cleared.")
            
            # select all items available
            items = [item.Name for item in pivot_field.PivotItems()]
            print(f"Found {len(items)} items in the pivot field '{product_field_caption}':")
            # print all the items founds (optional) 

            # only select the product to filter
            selected_items = [i for i in items if i == product_to_select]
            if not selected_items:
                raise Exception("No items found error.")
            
            # display the selected (filter items) only non-blank values
            pivot_field.VisibleItemsList = selected_items
            
            print(f"✅ Dates updated successfully.")
        else:
            print(f"Pivot Table '{pivot_table_name}' not found on the sheet.")

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
    product_field_caption = "Product"
    product_to_select = "[TotalMarket].[Product].&[NESTLE 100 GRAND BAR 42G]"

    select_single_product(filename, sheet_name, pivot_table_name, product_field_caption, product_to_select)