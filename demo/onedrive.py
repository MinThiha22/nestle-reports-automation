from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from dotenv import load_dotenv
import time,os

def click_button(driver, type):
  try:
    button = WebDriverWait(driver, 10).until(
      EC.element_to_be_clickable((By.CSS_SELECTOR, f"button[data-testid='{type}']"))
    )
    button.click()
    print(f"Clicked {type}")
    return True
  except Exception as e:
    print(f"Error: {e}")
    return False
  
def dismiss_popup(driver):
  try:
    no_thanks = WebDriverWait(driver, 2).until(
        EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'No thanks')]"))
    )
    no_thanks.click()
    time.sleep(1)
    print("Popup dismissed")
  except TimeoutException:
      pass

def login_onedrive():
  profile_path = os.path.expanduser("~/Library/Application Support/Microsoft Edge/Default")
  edge_options = Options()
  edge_options.add_argument(f"user-data-dir={profile_path}")
  edge_options.add_argument("--profile-directory=Default") 

  service = Service('./service/msedgedriver')
  driver = webdriver.Edge(service=service, options=edge_options)

  load_dotenv()
  email = os.getenv("EMAIL")
  password = os.getenv("PASSWORD")

  try:
    driver.get("https://onedrive.live.com/login")
    dismiss_popup(driver)
    # Front Page
    WebDriverWait(driver, 10).until(
      EC.frame_to_be_available_and_switch_to_it((By.TAG_NAME, "iframe"))
    )
    email_field = WebDriverWait(driver, 10).until(
      EC.presence_of_element_located((By.ID, "emailTextInput")))
    email_field.send_keys(email)
    print("Email entered")

    next_button = WebDriverWait(driver, 10).until(
      EC.element_to_be_clickable((By.ID, "nextButton")))
    next_button.click()
    driver.switch_to.default_content()
    dismiss_popup(driver)
    
    # Another username entry
    try:
      username_field = WebDriverWait(driver, 10).until(
      EC.presence_of_element_located((By.ID, "usernameEntry")))
      username_field.send_keys(email)
      click_button(driver, "primaryButton")
    except TimeoutException:
      print("Skipped another username entry")

    password_field = WebDriverWait(driver, 10).until(
      EC.presence_of_element_located((By.ID, "passwordEntry")))
    password_field.send_keys(password)
    click_button(driver, "primaryButton")
  
    dismiss_popup(driver)
    click_button(driver, "secondaryButton")
    
  except Exception as e:
    print(f"Error: {e}")

  finally:
    time.sleep(60) # 1 minutes
    driver.quit()

if __name__ == "__main__":
  login_onedrive()

  