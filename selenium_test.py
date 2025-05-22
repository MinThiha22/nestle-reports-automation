from selenium import webdriver
from selenium.webdriver.edge.service import Service
import time

service = Service('./service/msedgedriver')
driver = webdriver.Edge(service=service)

driver.get("https://google.com")
time.sleep(2)
driver.quit()