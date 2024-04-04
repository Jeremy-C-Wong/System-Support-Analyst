# Script to clean up NON_FILE_CATALOG_REQUEST spreadsheet
# to reduce amount of manual cleaning

# Abbreviations:
# - nfcr = Non-File Catalog Request

# Prerequisites:
# - Python interpreter installed on system
# - pip install pandas
# - pip install xlrd
# - pip install openpyxl

# - Rename the Non-File Catalog Request file to nfcr.xls

import numpy as np
import pandas as pd
from collections import namedtuple
import os

