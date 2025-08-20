from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
import os

downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
os.makedirs(downloads_folder, exist_ok=True)

def export_file(page):
    '''Export the file from the page.'''
    try:
      print("=== Starting export process... ===")
      print("=== Export Custom Attributes... === ")
      # Navigate to the "Custom Attributes" node     
      page.locator("#url_customattributes").click()
      page.locator("text=Export / Import").click()
      page.locator("#import_export_actions").click()
      page.locator("a:has(span:text('Export custom attributes'))").click()

      # Select groups and export the file
      select_group_selection_for_power_bi(page)

      print("=== Export process completed successfully. ===")
      print("=== Export Standart Attributes... ===")
      # Navigate to the "Custom Attributes" node
      page.locator("#import_export_actions").click()
      page.locator("a:has(span:text('Export standard attributes'))").click()
            
      # Select groups and export the file
      select_group_selection_for_power_bi(page)
      
      # Wait for the export to complete
      page.wait_for_timeout(3000)  # Wait for 3 seconds to ensure the export is processed
      # Download the file
      select_file_for_download(page)
      
    except Exception as e:
      print(f"Error locating and clicking on the export button: {e}")

      
'''Selecet Grooup Selection for Power BI and export the file'''
def select_group_selection_for_power_bi(page):
    # select the Category Hierarchy dropdown
    try:
      page.locator("button.form-control:has-text('Merch Category by Brand')").click()
      page.locator("li.ng-binding:has-text('Category Hierarchy')").click()
      page.locator("span.dynatree-node:has(a.dynatree-title:text('Custom Groups'))").click()
      page.locator("span.dynatree-node:has(a.dynatree-title:text('Favorites'))").click()
      page.locator("a.dynatree-title:text('Group Selection for Power BI')").click()
      page.locator("button.btn-primary:text('Export List')").click()
      page.wait_for_timeout(3000)  # Wait 3 sec before hit cancel
      page.locator("button.btn-default[ng-click=\"$dismiss('Cancelled')\"]").click()
    except PlaywrightTimeoutError as e:
      print(f"Timeout error while locating elements: {e}")
      
def select_file_for_download(page):
    page.wait_for_timeout(10000)  # Wait for 10 seconds to ensure the files are ready for download
    try:
      page.locator("a:has(span#dh-header-inbox)").click()  # Click on the message center icon
      '''Wait for the page to load completely'''
      page.wait_for_load_state('networkidle', timeout=30000) # Increased to 30 seconds for safety
      # 🔁 Refresh the page to ensure new messages appear
      page.wait_for_timeout(10000)  # wait 10 sec after reload, to ensure the messages are loaded
      page.reload(wait_until="networkidle") # reload the page to ensure the latest messages are loaded
      
      # Wait for the iframe with file to download appear
      '''Message center are in a iframe, so we need to switch to it'''
      iframe_element = page.wait_for_selector("iframe#messages-frame", timeout=10000)
      iframe = iframe_element.content_frame()
      iframe.wait_for_selector("div.message-link", timeout=10000)

      # Click the first message link
      message_links = iframe.locator("div.message-link")
      count = message_links.count() # Get the count of message links

      if count > 0:
          message_links.nth(0).click()
          print("✅ Clicked first message link.")
          download_file(page) # Call the download function to handle the file download
          
          page.wait_for_timeout(2000)  # Wait for 2 seconds to ensure the message is loaded
          message_links.nth(1).click()
          print("✅ Clicked second message link.")
          download_file(page) # Call the download function to handle the file download
      else:
          print("❌ No message links found.")
      # go back to the main page   
      page.locator("a#url_myworkspace").click()   

    except PlaywrightTimeoutError as e:
      print(f"Error during file download: {e}")        

def download_file(page):
  # Wait for the download to finish
      frame = page.frame(name="messages-frame")
      with page.expect_download(timeout=30000) as download_info: # this waits for the download to finish
        download_button = frame.locator("a:has(span.icon-download)").first
        download_button.wait_for(state="visible", timeout=10000)
        download_button.click()
      download = download_info.value
      download.path() 
      filename = download.suggested_filename # file name from the download
      # Save in Windows Downloads folder with suggested filename
      final_path = os.path.join(downloads_folder, filename)
      download.save_as(final_path)
      print("✅ Download finished. Moving on...")
      
      # Close the message modal
      frame.wait_for_selector("a#message-modal_modalCloseX", timeout=10000)
      frame.locator("a#message-modal_modalCloseX").click()