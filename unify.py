import os, time
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
import time
import re
from datetime import datetime, timedelta

DOWNLOADS_PATH = os.path.join(os.path.expanduser("~"), "Downloads")
def sanitize_filename(name):
  return re.sub(r'[\\/*?:"<>|]', "_", name)
def login_unify(p):
  load_dotenv()
  username = os.getenv("UNIFY_USERNAME")
  password = os.getenv("UNIFY_PASSWORD")
  
  context = p.chromium.launch_persistent_context(
    user_data_dir="./edge_profile",
    headless=False,
    channel="msedge",
    accept_downloads=True,
    downloads_path=DOWNLOADS_PATH
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

def navigate_export(page):
  try:
    # Navigate to Favourite, Choose Flat files 2
    locate_and_action(page, '#FavoritesLink', description="Favourite")
    locate_and_action(page, 'span', has_text="Flat Files 2", description="Flat Files 2")
    locate_and_action(page, 'div.thumb-box', has_text="Flat File - CD", description="CD Card")
    time.sleep(2)
    
    export_flat_file(page, "Flat File - CD")
    export_flat_file(page, "Flat File - NWNI")
    export_flat_file(page, "Flat File - PSNI")
    export_flat_file(page, "Flat File - FSSI")
    export_flat_file(page, "Flat File - Petrol CDNISI")
    
    # Navigate to Favourite again, Choose Flat File - TSM NI/SI
    locate_and_action(page, '#FavoritesLink', description="Favourite")
    locate_and_action(page, 'span', has_text="Flat File - TSM NI/SI", description="Flat File - TSM NI/SI")
    export_flat_file(page,"Flat File - TSM NI/SI")
    
    # Navigate to Favourite againk, Choose Flat File - Chemist Warehouse
    locate_and_action(page, '#FavoritesLink', description="Favourite")
    locate_and_action(page, 'span', has_text="Flat File - Chemist Warehouse", description="Flat File - Chemist Warehouse")
    export_flat_file(page, "Flat File - Chemist Warehouse")
        
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

import time

def wait_for_notification_download(page, noti_count=7, timeout_ms=10800000, check_interval=30):
  print(f"🔔 Waiting for notification count to reach {noti_count} (timeout: {timeout_ms // 60000} min)")
  noti_span = page.locator('a.fa-bell span.notification')
  start = time.time()

  while (time.time() - start) * 1000 < timeout_ms:
    try:
      noti_span.wait_for(state='attached', timeout=5000)
      count_text = noti_span.inner_text().strip()
      count = int(count_text) if count_text.isdigit() else 0
      print(f"🔎 Current count: {count}")
      if count == noti_count:
        print("✅ Target count reached! Clicking notification.")
        download_from_notifications(page)
        return
    except Exception as e:
      print(f"⚠️ Error: {e}")
    time.sleep(check_interval)
  raise TimeoutError(f"⏰ Timeout: Notification count did not reach {noti_count} in time.")

def parse_notification_time(time_str):
  now = datetime.now()
  if "Today" in time_str:
    time_part = time_str.replace("Today, ", "")
    time_obj = datetime.strptime(time_part, "%I:%M %p")
    return now.replace(hour=time_obj.hour, minute=time_obj.minute, second=0, microsecond=0)
  elif "Yesterday" in time_str:
    time_part = time_str.replace("Yesterday, ", "")
    time_obj = datetime.strptime(time_part, "%I:%M %p")
    yesterday = now - timedelta(days=1)
    return yesterday.replace(hour=time_obj.hour, minute=time_obj.minute, second=0, microsecond=0)
  else:
    return None

def get_notification_items(page):
  notifications = []
  locate_and_action(page,'a.fa-bell',description='Noti Bell Icon')
  try:
    page.wait_for_selector('ul.notifications-list li', timeout=5000)
    notification_elements = page.locator('ul.notifications-list li').all()
      
    for element in notification_elements:
      try: 
        title_element = element.locator('p.notify-title')
        file_name = title_element.get_attribute('title') or title_element.inner_text()
        
        # Extract time
        time_element = element.locator('p.notify-date')
        time_str = time_element.inner_text()
        
        # Extract status
        status_element = element.locator('span.notify-status')
        status = status_element.inner_text()
        
        notifications.append({
            'element': element,
            'file_name': file_name,
            'time_str': time_str,
            'time_obj': parse_notification_time(time_str),
            'status': status
        })
      except Exception as e:
        print(f"Error parsing notification item: {e}")
        continue
  except Exception as e:
      print(f"Error getting notifications: {e}")
  print(notifications) 
  locate_and_action(page,'a.fa-bell',description='Noti Bell Icon')
  return notifications

def download_from_notifications(page): 
  target_files = ["Flat File - CD", "Flat File - NWNI", "Flat File - PSNI", "Flat File - FSSI", "Flat File - Petrol CDNISI", "Flat File - Chemist Warehouse", "Flat File - TSM NI/SI"]
  notifications = get_notification_items(page)
  if not notifications:
    print("❌ No notifications found")
    return False
  valid_notifications = []
  downloaded_files = set()
    
  for notif in notifications:
    if (notif['status'] == 'Export Complete' and any(target_file in notif['file_name'] for target_file in target_files)):
      file_type = None
      for target in target_files:
        if target in notif['file_name']:
          file_type = target
          break
        
      if file_type and file_type not in downloaded_files:
        valid_notifications.append(notif)
        downloaded_files.add(file_type)
        if len(downloaded_files) >= 7:
          break
    
  print(f"🎯 Found {len(valid_notifications)} valid notifications to download")

  download_count = 0
  for notif in valid_notifications[:7]:  # Limit to 7 downloads
    try:
      locate_and_action(page,'a.fa-bell',description='Noti Bell Icon')
      with page.expect_download(timeout=600000) as download_info:
        notif['element'].click()
        download = download_info.value
        safe_file_name = sanitize_filename(notif['file_name']) + ".xlsx"
        download.save_as(os.path.join(DOWNLOADS_PATH, safe_file_name))
        print(f"✅ Saved: {safe_file_name}")
        download_count += 1
        time.sleep(2)  # Optional: wait a bit between downloads
    except Exception as e:
      print(f"❌ Error downloading {notif['file_name']}: {e}")
      
  print(f"✅ Successfully initiated {download_count} downloads from notifications")
  return download_count > 0

def unify_automation():
  p = sync_playwright().start()     
  page, context = login_unify(p)
  get_notification_items(page)
  return p, context

def close_automation(p,context):
  if p:
    p.stop()
  if context:
    context.close()

if __name__ == "__main__":
  with sync_playwright() as p:
    page, context = login_unify(p)
    navigate_export(page)
    download_from_notifications(page)                   
    input("Press Enter to exit and close browser...")
    context.close()
    