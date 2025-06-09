from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
import time

service = Service('./service/msedgedriver')
driver = webdriver.Edge(service=service)

driver.get("https://the-internet.herokuapp.com/login")

# Fill in username and password
driver.find_element(By.ID, "username").send_keys("tomsmith")
driver.find_element(By.ID, "password").send_keys("SuperSecretPassword!")

# Click login
driver.find_element(By.CSS_SELECTOR, "button.radius").click()

time.sleep(5)

# Print result message
message = driver.find_element(By.ID, "flash").text
print(message)
time.sleep(5)
driver.quit()