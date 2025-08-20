from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time

# Set up browser
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")  # or use --headless for silent login
driver = webdriver.Chrome(service=Service(
    ChromeDriverManager().install()), options=options)

# Open Circana login page
driver.get("https://your.circana.login.url/")  # replace with actual login URL

# Wait for Microsoft redirect and login form to load
time.sleep(3)  # You should ideally use WebDriverWait

# Enter email
driver.find_element(By.NAME, "loginfmt").send_keys("your_email@domain.com")
driver.find_element(By.ID, "idSIButton9").click()

time.sleep(3)

# Enter password
driver.find_element(By.NAME, "passwd").send_keys("your_password")
driver.find_element(By.ID, "idSIButton9").click()

time.sleep(3)

# Optional: Handle "Stay signed in?" prompt
try:
    # Or "idSIButton9" to stay signed in
    driver.find_element(By.ID, "idBtn_Back").click()
except:
    pass

# Done — you can now access protected pages
print("Logged in!")
