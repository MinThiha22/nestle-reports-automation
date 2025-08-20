from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
import os

class ImportResult:
    def __init__(self, success=False, has_error=False, message=""):
        self.success = success
        self.has_error = has_error
        self.message = message

def import_file(page, file_path=None):
  '''Import file to the page.'''
  try:
    print("=== Starting export process... ===")
    print("=== Import File... === ")
    
    # navigate to the import page
    page.locator("#url_customattributes").click()
    page.locator("text=Export / Import").click()
    page.locator("#import_export_actions").click()
    page.locator("a:has(span:text('Import custom attributes'))").click()
   
    '''inject the file to the site and click on the Import button'''
    # upload the file and start the import process
    page.set_input_files('input[name="file"]', file_path)
    page.click('button:has-text("Import Data")')
   # Wait for import to process
    page.wait_for_load_state('networkidle', timeout=30000)
    
    # Check import result
    result = accept_file(page)
    '''check if the file went through or not'''
    if result.success:
      page.reload()
      print("=== ✅ Import process completed successfully. ===")
      expand_if_needed(page)
      resubmit_report(page)
    else:
        print("=== ❌ Import process failed. Please check the logs for errors. ===")
        return False    
  
  except PlaywrightTimeoutError as e:
    print(f"Timeout error while locating elements: {e}")     

def accept_file(page):
  '''Accept the file after import and check the status.'''
  try:
    status_row = page.locator("table tbody tr").nth(0) # select the first row of the table
    status = status_row.locator("td").nth(2).text_content()  # Get the text content of the third cell
    status_error = status_row.locator("td").nth(5).text_content() # check if there is an error in the import process
    
    if status == "PENDING":
        if status_error == "0":
            print("✅ No errors found in the import process.")
            accept_btn = page.locator("span[uib-tooltip='Accept']")
            accept_btn.wait_for(state="visible")
            accept_btn.click()
            page.wait_for_timeout(2000)
            yes_btn = page.locator('button[ng-click="yes()"]')
            yes_btn.wait_for(state="visible")
            yes_btn.click()
            page.wait_for_timeout(2000)
            accept_btn = page.locator("button:has-text('Ok')")
            accept_btn.wait_for(state="visible")
            accept_btn.click()
            return ImportResult(True, False, "Import accepted successfully")
        else:
            print("⚠️ Errors found in the import process.")
            reject_btn = page.locator("span[uib-tooltip='Reject']")
            reject_btn.wait_for(state="visible")
            reject_btn.click()
            page.wait_for_timeout(2000)
            yes_btn = page.locator('button[ng-click="yes()"]')
            yes_btn.wait_for(state="visible")
            yes_btn.click()
            page.wait_for_timeout(2000)
            accept_btn = page.locator("button:has-text('Ok')")
            accept_btn.wait_for(state="visible")
            accept_btn.click()
            return ImportResult(False, True, "Import rejected due to errors")
    elif status == "REJECTED":
      print("❌ Import process was rejected. Please check the logs for errors.")
      return ImportResult(False, True, "Import rejected due to errors")
    else:
        print("Page is getting refreshed, please wait...")
        page.wait_for_timeout(3000)  # Wait for the page to refresh
        page.reload()
        page.wait_for_load_state('networkidle', timeout=30000)  # Wait for the page to load completely
        accept_file(page)  # Call the function again to check the status after reload
  except PlaywrightTimeoutError as e:
    print(f"Error while checking import status: {e}")
    return False  

# Expand the dynatree folders if needed  
def expand_if_needed(page):
  try:
    print("=== Starting download process... ===")
    page.click("#url_reports>>span:text('Reports')")
  
    """Expand only the needed dynatree folders by checking dynatree-expanded class."""
    
    for label in ["Shared", "FNZ Nestle New Zealand Limited", "9. Store Level"]:
      node = page.locator(f"span.dynatree-node:has(span.dynatree-expander):has(a.dynatree-title:has-text('{label}'))")
      class_attr = node.get_attribute("class") or ""
    
      if "dynatree-expanded" not in class_attr:
          print(f"📂 Expanding: {node}")
          expander = node.locator("span.dynatree-expander")
          expander.click()
          page.wait_for_timeout(500)
      else:
          print(f"✅ Already expanded: {node}")

    # Finally, click F&B
    FnB = page.get_by_role("link", name="F&B", exact=True)
    if FnB.is_visible():
        FnB.click()
        print("🎯 Clicked F&B")
    else:
        raise Exception("❌ F&B is not visible after expanding.")
      
  except PlaywrightTimeoutError as e:
    print(f"Timeout error while locating elements: {e}")    

# Resubmit the report after import    
def resubmit_report(page):
  '''Resubmit the report.'''
  print("=== Starting resubmit process... ===")
  page.locator(".btn.dropdown-toggle.undraggable").first.click()
  page.locator(".context-menu-text").first.click()
  page.wait_for_timeout(2000)  # Wait for 2 seconds to ensure the
  page.locator("#btnSubmit").first.click()
  

    
    

