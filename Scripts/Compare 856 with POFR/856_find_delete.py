# Script to find Unprocessed 856's that do not have a corresponding
# PO Number in Purchase Orders for Receiving

# Abbreviations:
# - usn = Unprocessed Shippping Notices
# - pofr = Purchase Orders for Receiving

# Prerequisites:
# - Python interpreter installed on system
# - pip install pandas
# - pip install xlrd
# - pip install openpyxl
# - pip install selenium

# - Rename the Purchase Orders for Receiving file to pofr.xlsx
# - Rename the Unprocessed Shipping Notices file to usn.xlsx

import numpy as np
import pandas as pd
from collections import namedtuple
import os
from dotenv import load_dotenv 

import time
# (Ensure you also have these imports at the top of your file if not already present)
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# Import Selenium modules
from selenium import webdriver
from selenium.webdriver.edge.options import Options


# Define file paths that Excel files will be read from
abs_path = os.path.dirname(__file__)

rel_usn_path = 'usn.xlsx'
rel_pofr_path = 'pofr.xlsx'

usn_path = os.path.join(abs_path, rel_usn_path)
pofr_path = os.path.join(abs_path, rel_pofr_path)

# Read in data from Excel files and put in Pandas dataframes
usn_df = pd.read_excel(usn_path)
pofr_df = pd.read_excel(pofr_path)

# Create Series objects from PO No columns in dataframes and cast to strs
usn_ser = usn_df['PO No']
pofr_ser = pofr_df['PO No']
usn_ser = usn_ser.astype('str')
pofr_ser = pofr_ser.astype('str')

# Gets POs that are marked as Fully Received
pofr_ser_y = pofr_df[pofr_df['Fully Received'] == 'Yes']
pofr_ser_y = pofr_ser_y.astype('str')

# Convert Series objects to numpy arrays
usn_np = usn_ser.to_numpy()
pofr_np = pofr_ser.to_numpy()
pofr_np_y = pofr_ser_y.to_numpy()

# Creates new numpy array with unique elements in usn_np not in pofr_np
nonexist_np = np.setdiff1d(usn_np, pofr_np)

# Find the fully received POs in the unprocessed shipping notices
fully_rcvd_in_usn = np.intersect1d(usn_np, pofr_np_y)

# Combines the nonexistent POs with the fully received POs
nonexist_np_all = np.union1d(nonexist_np, fully_rcvd_in_usn)

# Removes POs with letters in them (not included in POFR spreadsheet)
nonexist_l = np.array([i for i in nonexist_np_all if i.isnumeric()])

usn_asset = usn_df[usn_df['Org/Asset Location'].isin(['1/1', '1/3', '1/5', '1/6'])]
usn_dup = usn_df[usn_df['Import Status'] == 'Duplicate 856']
usn_how = usn_df[usn_df['Vendor Name'] == 'HOWMEDICA OSTEONICS CORP']
usn_del_df = pd.concat([usn_asset, usn_dup, usn_how], axis=0)
usn_del_ser = usn_del_df['PO No']
usn_del_ser = usn_del_ser.astype('str')

po_numbers = np.concatenate((nonexist_l, usn_del_ser), axis=0)
po_numbers = np.unique(po_numbers)


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

delay = 30   # three seconds before timeout
try:
    pass_field = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.NAME, "Enter your Password")))
    pass_field.send_keys(PC_PASSWORD)
    pass_field.send_keys(Keys.RETURN)
except TimeoutException as e:
    print(e)

try:
    erp_but = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.LINK_TEXT, "ERP / Materials Management")))
    erp_but.click()
except TimeoutException as e:
    print(e)

WebDriverWait(driver, delay).until(lambda d: len(d.window_handles) > 1)
driver.switch_to.window(driver.window_handles[-1])
print("Switched to new tab")
print(driver.title)


try:
    # 1. Wait for the 'frLeft' frame to load and switch context into it
    WebDriverWait(driver, 15).until(
        EC.frame_to_be_available_and_switch_to_it((By.NAME, "frLeft"))
    )

    # 2. Wait for the button to be clickable inside that frame
    # Using contains(text(), ...) handles the extra text and image tag inside the <td>
    work_in_materials_btn = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.XPATH, "//td[contains(@class, 'TOCMenu') and contains(text(), 'Work in Materials')]"))
    )

    # 3. Click the button
    work_in_materials_btn.click()
    print("Successfully clicked 'Work in Materials'!")

finally:
    # 4. If you need to interact with elements in other frames later, 
    # you must first return to the main top-level document structure:
    driver.switch_to.default_content()


try:
    # 1. Wait for the 'frLeft' frame to load and switch context into it
    WebDriverWait(driver, 15).until(
        EC.frame_to_be_available_and_switch_to_it((By.NAME, "frLeft"))
    )

    # 2. Wait for the button to be clickable inside that frame
    # Using contains(text(), ...) handles the extra text and image tag inside the <td>
    receiving_btn = WebDriverWait(driver, delay).until(
        EC.element_to_be_clickable((By.XPATH, "//td[contains(@class, 'TOCMenu') and contains(text(), 'Receiving')]"))
    )

    # 3. Click the button
    receiving_btn.click()
    print("Successfully clicked 'Receiving'!")

finally:
    # 4. If you need to interact with elements in other frames later, 
    # you must first return to the main top-level document structure:
    driver.switch_to.default_content()


try:
    # 1. Wait for the 'frLeft' frame to load and switch context into it
    WebDriverWait(driver, delay).until(
        EC.frame_to_be_available_and_switch_to_it((By.NAME, "frLeft"))
    )

    # 2. Wait for the button to be clickable inside that frame
    # Using contains(text(), ...) handles the extra text and image tag inside the <td>
    usn_btn = WebDriverWait(driver, delay).until(
        EC.element_to_be_clickable((By.XPATH, "//td[contains(@class, 'TOCAction') and contains(text(), 'Unprocessed Advanced Ship Notices 856')]"))
    )

    # 3. Click the button
    usn_btn.click()
    print("Successfully clicked 'Unprocessed Advanced Ship Notices 856'!")

finally:
    # 4. If you need to interact with elements in other frames later, 
    # you must first return to the main top-level document structure:
    driver.switch_to.default_content()

WebDriverWait(driver, delay).until(lambda d: len(d.window_handles) > 2)
driver.switch_to.window(driver.window_handles[-1])
print("Switched to new tab")
print(driver.title)

delete_sum = 0

# 2. Iterate through each PO and delete the corresponding row(s)
for po in nonexist_l:
    print(f"--- Processing PO: {po} ---")
    try:
        # A. Locate the "PO No" text field by its ID, clear it, type the PO, and hit Enter
        po_input = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.ID, "qf_PONo")) 
        )
        po_input.clear()
        po_input.send_keys(po)
        po_input.send_keys(Keys.RETURN)
        
        # Give the ERP system a moment to reload/filter the table initially
        time.sleep(5)

        rows_deleted = 0
        
        # Continuously look for rows to delete until none are left
        while True:
            try:
                # Check if there are any "Menu" links visible on the page.
                # Using a short 3-second wait. If it times out, we assume there are no more rows.
                menu_btns = WebDriverWait(driver, 5).until(
                    EC.presence_of_all_elements_located((By.LINK_TEXT, "Menu"))
                )
            except TimeoutException:
                # No more "Menu" links found. Break out of the while loop and move to the next PO.
                break
            
            # B. Click the "Menu" link on the first available row in the list
            menu_btn = menu_btns[0]
            menu_btn.click()
            
            # C. Click "Delete" from the popup menu
            delete_btn = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.XPATH, "//a[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'delete')] | //div[contains(@class, 'Menu')]//td[contains(text(), 'Delete')]"))
            )
            delete_btn.click()
            
            # D. Handle the "Are You Sure..." Webpage Dialog using the iframe
            # 1. Wait for the iframe to load and switch into it
            WebDriverWait(driver, 15).until(
                EC.frame_to_be_available_and_switch_to_it((By.ID, "dialog-body"))
            )

            # 2. Wait for the dialog's Submit button to appear and click it using its ID
            submit_btn = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.ID, "butOK"))
            )
            submit_btn.click()
            
            # 3. Switch back to the main page context so the next loop iteration can find the PO text box
            driver.switch_to.default_content()
            
            rows_deleted += 1
            delete_sum += 1
            print(f"Deleted row {rows_deleted} for PO: {po}")
            
            # Brief pause to let the deletion process and page refresh before checking for the next row
            time.sleep(3)

        if rows_deleted == 0:
            print(f"No rows found to delete for PO: {po}")
        else:
            print(f"Successfully finished deleting all {rows_deleted} row(s) for PO: {po}")

    except TimeoutException:
        print(f"Timeout error for PO {po}. The PO search box or a required element didn't load in time.")
    except Exception as e:
        print(f"An unexpected error occurred processing PO {po}: {e}")

print("Finished processing all PO numbers.")
print(f"Deleted {delete_sum} rows.")
