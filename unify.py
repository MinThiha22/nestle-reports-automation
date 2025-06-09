import os, time
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
import time

def login_unify(p):
  load_dotenv()
  username = os.getenv("UNIFY_USERNAME")
  password = os.getenv("UNIFY_PASSWORD")
  
  context = p.chromium.launch_persistent_context(
    user_data_dir="./edge_profile",
    headless=False,
    channel="msedge"
  )
  page = context.pages[0]
  page.goto("https://unify.ap.iriworldwide.com/client1/index.html", wait_until="load")
  
  try:
    # If redirected to landing page, assume already logged in
    page.wait_for_url("https://unify.ap.iriworldwide.com/client1/plus/landing/0", timeout=1000)
    print("Already logged in. Redirected to landing page.")
  except PlaywrightTimeoutError:
    print("Login page detected. Proceeding with login...")
    
    try:
      page.wait_for_selector("#userID", timeout=5000)
      page.fill("#userID", username)
      page.fill("#password", password)
      page.click("#login")
      # Wait for landing page after login
      page.wait_for_url("https://unify.ap.iriworldwide.com/client1/plus/landing/0", timeout=300000)
      print("Login successful. Reached landing page.")
    except PlaywrightTimeoutError:
        print("Login failed or took too long.")
  return page, context

def automate_download(page):
  try:
    # Navigate to Favourite, Choose Flat files 2
    locate_and_action(page, '#FavoritesLink', description="Favourite")
    locate_and_action(page, 'span', has_text="Flat Files 2", description="Flat Files 2")
    locate_and_action(page, 'div.thumb-box', has_text="Flat File - CD", description="CD Card")
    time.sleep(2)
    
    export_flat_file(page, "Flat File - CD")
    export_flat_file(page, "Flat File - NWNI")
    export_flat_file(page, "Flat File - PSNI")
    export_flat_file(page, "Flat File - PSNI")
    export_flat_file(page, "Flat File - FSSI")
    export_flat_file(page, "Flat File - Petrol CDNISI")
  except Exception as e:
    print("Error during download automation:", e)
    
    
def export_flat_file(page, file_name):
  print(f"\n--- Exporting {file_name} ---")
  # Switch to file tab
  locate_and_action(page, 'ul#report-nav-scroll li.reportNavLi', has_text=file_name, description=f"{file_name} tab")
  
  # Click Action button and choose export
  locate_and_action(page, 'div.dashboard-action span.db-action-link', has_text="Action", description="Action Button")
  locate_and_action(page, '#reportContainer div.actionModal li.action-modal-item span', has_text="Export", description="Export Option")

  # Check SelectAll Box from both Geography and Time
  locate_and_action(page, 'div.selectAll label.check-label', action='Check', description="SelectAll box")
  time.sleep(2)
  locate_and_action(page, 'div.iterate-select select', action='Select', option='1: Object', description="Time Option")
  time.sleep(2)
  locate_and_action(page, 'div.selectAll label.check-label', action='Check', description="SelectAll box")
  time.sleep(2)
  
  # Select Excel and Choose Pivot Table
  locate_and_action(page, 'div.fileType label.check-label span', has_text="Excel Spreadsheet", action="Check", description="Excel file type")
  locate_and_action(page, 'ul li label.check-label span', has_text="Pivot Table", action="Check", description="Pivot Table")
  time.sleep(3)
  
  # Click Export
  locate_and_action(page, 'div.modal-footer div.exp-footer-button button', has_text="Export", description="Export button")
  time.sleep(3) 
  
  # Close Information modal
  locate_and_action(page, 'div.modal-dialog div.modal-content div button', has_text="Okay", description="Okay button")
  time.sleep(5) 

def locate_and_action(page, selector, has_text=None, action="Click", option=None, description="", timeout=10000):
  try:
    element = page.locator(selector, has_text=has_text) if has_text else page.locator(selector)
    element.wait_for(state="visible", timeout=timeout)
    if action == "Click":
      element.click()
    elif action == "Check":
      element.check()
    elif action == "Select":
      element.select_option(option)
    else:
      print(f"Unsupported Action")
    print(f"{action}ed {description}")
    return element
  except Exception as e:
    print(f"Error during {action} {description}: {e}")

def unify_automation():
  with sync_playwright() as p:
    page, context = login_unify(p)
    automate_download(page)
    return context

if __name__ == "__main__":
  with sync_playwright() as p:
    page, context = login_unify(p)
    automate_download(page)
    input("Press Enter to exit and close browser...")
    context.close()

    