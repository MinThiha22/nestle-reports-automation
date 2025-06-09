import os
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
import time

load_dotenv()
username = os.getenv("GITHUB_USERNAME")
password = os.getenv("GITHUB_PASSWORD")

def test_github_login():
    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir="./edge_profile",
            headless=False,
            channel="msedge"
        )
        page = context.pages[0]
        page.goto("https://github.com/login", wait_until="load")

        try:
          page.wait_for_url("https://github.com/", timeout=2000)
          print("Already log in")
        except PlaywrightTimeoutError:
          try:
            print("Login page detected. Proceeding with login...")
            print("Logging in to GitHub...")
            page.fill("#login_field", username)
            page.fill("#password", password)
            page.click("input[type=submit]")
            page.wait_for_url("https://github.com/")
            print("Logged in")
          except PlaywrightTimeoutError:
            print("Login failed or took too long.")
        time.sleep(10)
        context.close()

if __name__ == "__main__":
  test_github_login()