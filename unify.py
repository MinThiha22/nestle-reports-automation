from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from dotenv import load_dotenv
import time,os

def click_button(driver, id):
  try:
    button = WebDriverWait(driver, 10).until(
      EC.element_to_be_clickable((By.ID, f"{id}"))
    )
    button.click()
    print(f"Clicked {id}")
    return True
  except Exception as e:
    print(f"Error: {e}")
    return False

def login_unify():
  user_data_dir = os.path.expanduser("~/Library/Application Support/Microsoft Edge")
  profile_dir = "Default"  # or "Profile 1", etc.
  edge_options = Options()
  edge_options.add_argument(f"--user-data-dir={user_data_dir}")
  edge_options.add_argument(f"--profile-directory={profile_dir}")

  service = Service('./service/msedgedriver')
  driver = webdriver.Edge(service=service, options=edge_options)

  load_dotenv()
  username = os.getenv("UNIFY_USERNAME")
  password = os.getenv("UNIFY_PASSWORD")

  try:
    driver.get("https://unify.ap.iriworldwide.com/client1/index.html")

    # Index Page/ Login
    username_field = WebDriverWait(driver, 10).until(
      EC.presence_of_element_located((By.ID, "userID")))
    username_field.send_keys(username)
    print("Email entered")

    password_field = WebDriverWait(driver, 10).until(
      EC.presence_of_element_located((By.ID, "password")))
    password_field.send_keys(password)
    print("Password entered")
    
    click_button(driver, "login")
    
  except Exception as e:
    print(f"Error: {e}")

  finally:
    time.sleep(60) # 1 minutes
    driver.quit()

if __name__ == "__main__":
  login_unify()
  

  