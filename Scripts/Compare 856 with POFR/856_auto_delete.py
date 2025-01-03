import os
from dotenv import load_dotenv 

# Import Selenium modules
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.options import Options

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

load_dotenv()

PC_USERNAME = os.getenv("PC_USERNAME")
PC_PASSWORD = os.getenv("PC_PASSWORD")

edge_options = Options()
edge_options.add_experimental_option("detach", True)

# Create a webdriver object with Edge as the browser
driver = webdriver.Edge(options=edge_options)

driver.get("https://pincai.premierinc.com/")

driver.maximize_window()

email_field = driver.find_element(By.NAME, "username")
email_field.send_keys(PC_USERNAME)
email_field.send_keys(Keys.RETURN)

delay = 3   # three seconds before timeout
try:
    pass_field = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.NAME, "password")))
    pass_field.send_keys(PC_PASSWORD)
    pass_field.send_keys(Keys.RETURN)
except TimeoutException as e:
    print(e)

try:
    erp_but = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.LINK_TEXT, "ERP / Materials Management")))
    erp_but.click()
except TimeoutException as e:
    print(e)

#menu_but = driver.find_element(By.LINK_TEXT, "Menu")

#menu_but.click()