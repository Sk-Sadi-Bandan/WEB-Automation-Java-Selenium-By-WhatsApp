from selenium import webdriver
import selenium
from selenium.webdriver.common.by import By
import time


with open("phonenum.txt", "r",encoding='utf-8') as file:
    mobile_numbers = file.readlines()
print(mobile_numbers)
try:
    option = webdriver.ChromeOptions()
    option.add_argument("--disable-extensions")
    option.add_argument("--disable-infobars")
    # option.add_argument("--disable-gpu")
    option.add_argument("--no-sandbox")
    option.add_argument("--disable-dev-shm-usage")
    option.add_argument("--disable-gpu")
    option.add_argument(f'--user-data-dir=C:/Users/{user}/AppData/Local/Google/Chrome/User Data')
    option.add_argument('--profile-directory=Profile 2')
    driver = webdriver.Chrome(options=option)
    driver.get("https://web.whatsapp.com/")
    time.sleep(20)
    group_button= driver.find_element(By.XPATH, "//span[@data-icon='new-chat-outline']")
    group_button.click()
    time.sleep(5)
    create_group = driver.find_element(By.XPATH, "//div[@class='_3_9fX']")
    time.sleep(2)
    create_group.click()
    time.sleep(2)
    search_box= driver.find_element(By.XPATH, "//input[@placeholder='Search name or number']")
    for mobile_number in mobile_numbers:
        search_box.clear()
        mobile_number1 = str(mobile_number).replace("\n","")
        search_box.send_keys(f"{mobile_number1}\n")
        time.sleep(1)
    forward= driver.find_element(By.XPATH,"//span[@data-icon='arrow-forward']")
    forward.click()
    time.sleep(2)
    group_name= driver.find_element(By.XPATH, "//div[@title='Group Subject (Optional)']")
    group_name.send_keys("Test Group Subject")
    time.sleep(5)
    group_create = driver.find_element(By.XPATH, "//div[@aria-label='Create group']")
    group_create.click()
    time.sleep(20)
except (selenium.common.exceptions.WebDriverException, FileNotFoundError) as e:
    print(f"An error occurred: {e}")
# driver.quit()