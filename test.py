import os
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright
import time

def login_unify():
    load_dotenv()
    username = os.getenv("UNIFY_USERNAME")
    password = os.getenv("UNIFY_PASSWORD")

    # User data directory and profile for session persistence


    with sync_playwright() as p:
        # Launch Edge (Chromium-based) with user profile
        browser = p.chromium.launch(
            
            headless=False,
            channel="msedge"  # Use Microsoft Edge browser channel
        )
        page = browser.new_page()

        try:
            page.goto("https://unify.ap.iriworldwide.com/client1/index.html")

            # Wait for username field and fill
            page.wait_for_selector("#userID", timeout=10000)
            page.fill("#userID", username)
            print("Email entered")

            # Wait for password field and fill
            page.wait_for_selector("#password", timeout=10000)
            page.fill("#password", password)
            print("Password entered")

            # Click login button by ID
            #page.click("#login")
            #print("Clicked login")

        except Exception as e:
            print(f"Error: {e}")

        # Keep browser open for 10 minutes (600 seconds)
        time.sleep(1)

        # Navigate to blank page before closing context
        browser.close()

if __name__ == "__main__":
    login_unify()