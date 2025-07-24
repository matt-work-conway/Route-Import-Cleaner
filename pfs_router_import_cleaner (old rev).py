#pfs_router_report_compiner.py


import os
import numpy as np
import pandas as pd
from datetime import date
import time as time
import math as math
import openpyxl
import xlrd
from openpyxl.utils.dataframe import dataframe_to_rows
import re
from pathlib import Path
import win32com.client

from pandas.io.html import _remove_whitespace

def my_remove_whitespace(x):
    return x

pd.io.html._remove_whitespace = my_remove_whitespace


main_path = os.path.dirname(__file__)



#file_path = os.path.join(main_path, 'Indented BOM')

file_path = os.path.join(main_path, 'Route Imports')

#file_path = 'P:\\Conway\\chrome_crawler\\PFS_ROUTER_IMPORTS'

def parse_folder(file_path):
    paths = []
    for entry in os.scandir(file_path):
        paths.append(entry)

    return paths

def excel_to_csv(excel_filepath, csv_filepath):
    """
    Reads an Excel file and saves it as a CSV file.

    Args:
        excel_filepath (str): The path to the Excel file.
        csv_filepath (str): The path to save the CSV file.
    """
    try:
        df = pd.read_html(excel_filepath)[0]

        
        df = pd.DataFrame(df)
       
        char_to_find = ' '
        replacement_char = ''

        # Iterate over all cells in the DataFrame
        for row_index, row in df.iterrows():
            for col_index, value in row.items():
                if isinstance(value, str) and char_to_find in value:
                    df.loc[row_index, col_index] = value.replace(char_to_find, replacement_char)




        df.to_csv(csv_filepath, index = False)
        print(f"Successfully converted '{excel_filepath}' to '{csv_filepath}'")
    except FileNotFoundError:
        print(f"Error: Excel file '{excel_filepath}' not found.")
    except Exception as e:
        print(f"An error occurred: {e}")

for entry in parse_folder(file_path):
    entry_csv = str(entry)
    entry_csv = entry_csv.replace('.xls', '.csv')
    entry_csv = entry_csv.replace('>', '')
    entry_csv = entry_csv.replace('\'', '')
    entry_csv = entry_csv[9:].strip()
    entry_csv = entry_csv.replace('.xls', '.csv')
    entry_csv = os.path.join(file_path,entry_csv)
    excel_to_csv(entry, entry_csv)



