import os
import time
from datetime import datetime
from playwright.sync_api import sync_playwright
from importFile import expand_if_needed
from login import login_humby

downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
os.makedirs(downloads_folder, exist_ok=True)
user_data_dir = os.path.join(os.getcwd(), "user-data")

CHECK_INTERVAL_SECONDS = 1 * 60 * 60  # 1 hours

def file_available(page):
    row = page.locator("table tbody tr").nth(0) # have to be nth(), first() doesn't work.!!!
    cell = row.locator("td").nth(1) 
    name = "job"+cell.inner_text().strip()
    status_icon = row.locator('[ng-show="!!objstatus"]')
    status_text = status_icon.get_attribute("uib-tooltip") or ""

    print("Status tooltip:", status_text)

    if status_text.upper().startswith("COMPLETE"):
        before_files = set(os.listdir(downloads_folder))

        file_link = row.locator("td").nth(4)
        try:
            # 🔁 Try to capture download via Playwright event
            with page.expect_download(timeout=30000) as download_info:
                file_link.locator("a:has-text('Store Level Report')").click()
            download = download_info.value
            final_path = os.path.join(downloads_folder, name)
            download.save_as(final_path)
            print(f"✅ Downloaded and saved as: {final_path}")
            return final_path

        except Exception as e:
            # 🟡 Fallback: event missed, but download might have happened
            print(f"⚠️ Primary download failed: {e}")
            print("⏳ Waiting to detect download manually...")

            try:
                downloaded_path = wait_for_new_download(before_files,name, timeout=30)
                print(f"✅ File found in Downloads: {downloaded_path}")
                return downloaded_path
            except TimeoutError:
                print("❌ Download not detected after fallback wait.")
                return False

    return False

def wait_for_new_download(before_files, name, timeout=30):
    print("⏳ Waiting for new download to appear...")
    for _ in range(timeout):
        current_files = set(os.listdir(downloads_folder))
        new_files = current_files - before_files

        # Filter out temp/incomplete files only
        valid_files = [
            f for f in new_files
            if not f.endswith(('.crdownload', '.tmp', '.part'))
        ]

        if valid_files:
            downloaded_file = valid_files[0]
            downloaded_path = os.path.join(downloads_folder, downloaded_file)

            # Use the provided 'name' as the new filename (preserve extension if present)
            _, ext = os.path.splitext(downloaded_file)
            if not ext:
                ext = ".zip"  # Default extension if missing

            new_path = os.path.join(downloads_folder, name + ext)
            os.rename(downloaded_path, new_path)
            print(f"📁 Renamed {downloaded_file} to {os.path.basename(new_path)}")
            return new_path

        time.sleep(1)

    raise TimeoutError("Download did not appear in time.")


def run_check(context):
    page = login_humby(context)
    expand_if_needed(page)
    
    # Wait for page to load
    page.wait_for_load_state("networkidle")
    downloaded_path = file_available(page)

    # === Check if the file is available ===
    if downloaded_path:
        return downloaded_path
    else:
        print(f"❌ File not ready yet. Will retry later.")
        return False
     