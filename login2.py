from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

# Initialize the browser driver (you'll need to download the appropriate one)
driver = webdriver.Chrome()  # or Firefox(), etc.

# Open the login page
driver.get('https://unify.ap.iriworldwide.com/client1/index.html')

# Find the username and password fields and submit button
username_field = driver.find_element(By.NAME, 'username')
password_field = driver.find_element(By.NAME, 'password')
submit_button = driver.find_element(By.XPATH, '//button[@type="submit"]')

# Enter credentials and submit
username_field.send_keys('anzsl610')
password_field.send_keys('Kashmir77$')
submit_button.click()

# Check if login was successful
if "Dashboard" in driver.title:
    print("Login successful!")
else:
    print("Login failed")

# Continue with other actions...
driver.quit()  # Close the browser when done