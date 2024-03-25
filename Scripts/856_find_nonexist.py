# Script to find Unprocessed 856's that do not have a corresponding
# PO Number in Purchase Orders for Receiving

# Abbreviations:
# - usn = Unprocessed Shippping Notices
# - pofr = Purchase Orders for Receiving

# Prerequisites:
# - pip install pandas
# - pip install xlrd
# - pip install openpyxl

# - Rename the Purchase Orders for Receiving file to pofr.xls
# - Rename the Unprocessed Shipping Notices file to usn.xls

import numpy as np
import pandas as pd
from collections import namedtuple
import os

# Define file paths that Excel files will be read from
abs_path = os.path.dirname(__file__)

rel_usn_path = 'Compare 856 with POFR/usn.xls'
rel_pofr_path = 'Compare 856 with POFR/pofr.xls'

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
nonexist_l = [i for i in nonexist_np_all if i.isnumeric()]

# Exports PO list to Excel file in same directory as other Excel files
nonexist_df = pd.DataFrame(nonexist_l)

# nonexist_df.to_excel('C:/Users/wongj/OneDrive/Documents/School/Supply Chain System Analyst/Scripts/Compare 856 with POFR/nonexist.xlsx')
nonexist_df.to_excel(os.path.join(abs_path, 'Compare 856 with POFR/nonexist.xlsx'))
