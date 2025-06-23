# Script to find Unprocessed 856's that do not have a corresponding
# PO Number in Purchase Orders for Receiving

# Abbreviations:
# - edice = EDI Confirm Exceptions

# Prerequisites:
# - Python interpreter installed on system
# - pip install pandas
# - pip install xlrd
# - pip install openpyxl
# = pip install monthdelta

# - Rename the EDI Confirm Exceptions file to edice.xls

import numpy as np
import pandas as pd
from collections import namedtuple
from datetime import datetime
from monthdelta import monthdelta
import os

# Define constants
MIN_ORDER_COST = 0
MIN_CONFIRM_COST = 0
RATIO_UPPER_TOL = 1.1
RATIO_LOWER_TOL = 0.75
MONTHS_BACK_RANGE = 2

# Define file paths that Excel files will be read from
abs_path = os.path.dirname(__file__)

rel_edice_path = 'edice.xls'

edice_path = os.path.join(abs_path, rel_edice_path)

# Read in data from Excel files and put in Pandas dataframes
edice_df = pd.read_excel(edice_path)

# Remove lines without Price Discrepancies, with Contract Nos, and where In Tolerance == Yes
edice_pd_df = edice_df[edice_df["Price Discrepancy"] == "Yes"]
edice_pd_df = edice_pd_df[edice_pd_df["Contract No"].isnull()]
edice_pd_df = edice_pd_df[edice_pd_df["In Tolerance "].isnull()]

# Remove lines with Ordered Cost <= 0 and Confirm Cost <= 0
edice_pd_df = edice_pd_df[edice_pd_df["Ordered Cost"] > MIN_ORDER_COST]
edice_pd_df = edice_pd_df[edice_pd_df["Confirm Cost"] > MIN_CONFIRM_COST]

# Calculate ratio of Confirm Cost to Ordered Cost and filter lines outside of thresholds
edice_pd_df["Cost Ratio"] = edice_pd_df["Confirm Cost"].div(edice_pd_df["Ordered Cost"])
edice_pd_df = edice_pd_df[edice_pd_df["Cost Ratio"] < RATIO_UPPER_TOL]  # upper bound
edice_pd_df = edice_pd_df[edice_pd_df["Cost Ratio"] > RATIO_LOWER_TOL]  # lower bound

# Remove lines with dates earlier than a certain point beyond the present
edice_pd_df["Acknowledge Date"] = pd.to_datetime(edice_pd_df["Acknowledge Date"])
now = datetime.now()
months_back = now - monthdelta(MONTHS_BACK_RANGE)
edice_pd_df = edice_pd_df[edice_pd_df["Acknowledge Date"] >= months_back]

# Filter out unnecessary columns
keep_cols = [
                'PO No',
                'PO Item No',
                'PO Item Description',
                'Confirm Cost',
                'Ordered Cost',
                'Cost Ratio'
            ]

edice_pd_df = edice_pd_df[keep_cols]

edice_pd_df = edice_pd_df.sort_values(['PO No', 'PO Item No'])

edice_pd_df.to_excel(os.path.join(abs_path, 'in_thresholds.xlsx'))