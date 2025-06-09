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
    locate_and_action(page, '#FavoritesLink', description="Favourite")
    locate_and_action(page, 'span', has_text="Flat Files 2", description="Flat Files 2")
    locate_and_action(page, 'div.thumb-box', has_text="Flat File - CD", description="CD Card")
    time.sleep(2)
    
    locate_and_action(page, 'div.dashboard-action span.db-action-link', has_text="Action", description="Action Button")
    locate_and_action(page, '#reportContainer div.actionModal li.action-modal-item span', has_text="Export", description="Export Option")
    
    time.sleep(2)
    locate_and_action(page, 'div.selectAll label.check-label', action='Check', description="SelectAll box")
    time.sleep(2)
    locate_and_action(page, 'div.iterate-select select', action='Select', option='1: Object', description="Time Option")
    locate_and_action(page, 'div.selectAll label.check-label', action='Check', description="SelectAll box")
    time.sleep(2)
    locate_and_action(page, 'div.fileType label.check-label span', has_text="Excel Spreadsheet", action="Check", description="Excel file type")
    time.sleep(2)
    locate_and_action(page, 'ul li label.check-label span', has_text="Pivot Table", action="Check", description="Pivot Table")
    time.sleep(2)
    locate_and_action(page, 'div.modal-footer div.exp-footer-button button', has_text="Export", description="Export button")
    time.sleep(3)
    locate_and_action(page, 'div.modal-dialog div.modal-content div button', has_text="Okay", description="Okay button")
    time.sleep(3)
    
  except Exception as e:
    print("Error during download automation:", e)

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

if __name__ == "__main__":
  with sync_playwright() as p:
    page, context = login_unify(p)
    automate_download(page)
    input("Press Enter to exit and close browser...")
    context.close()

    