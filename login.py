import time
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
import os
from exportFile import export_file
from importFile import import_file


def login_humby(context):
  # Load environment variables
  load_dotenv()

  # Get credentials and login URL
  username = os.getenv("EMAIL")
  password = os.getenv("PASSWORD")
  login_url = os.getenv("LOGIN_URL")
  # logged_in_url = os.getenv("LOGGED_IN_URL")

  page = context.new_page()
  page.goto(login_url, wait_until="load")

  try:
      page.wait_for_selector("#userNameInput", timeout=5000)
      print("Login page detected. Proceeding with login...")
      page.fill("#userNameInput", "")  # Clear autofill
      page.fill("#userNameInput", username, timeout=3000)
      page.fill("#passwordInput", password, timeout=3000)
      page.click("#submitButton")

      # Wait for element on landing page instead of relying on URL
      page.wait_for_selector("text=Reports", timeout=15000)
      print("Login successful. Reached landing page.")

      return page

  except Exception as e:
      print(f"Error launching browser or navigating to login page: {e}")
      return None
    

def close_automation(p,context):
  if p:
    p.stop()
  if context:
    context.close()

if __name__ == "__main__":
  with sync_playwright() as p:
    browser = p.chromium.launch(
    headless=False,  # Set to False to see the browser UI
    args=["--start-maximized"]
    )
    context = browser.new_context(no_viewport=True)
    page = login_humby(context)   
    export_file(page)  # Call the download function   
    import_file(page) 
    time.sleep(10)  # Wait 
    context.close()
    browser.close()
