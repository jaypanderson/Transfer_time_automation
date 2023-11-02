"""
A script that will take the results from the transfer_time script and insert key data into a billing
form that will be delivered to the parents for extra charges incurred during that month.
"""

from __future__ import annotations
from typing import Union
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from datetime import datetime
import openpyxl
from openpyxl import Workbook
# from openpyxl import styles
from openpyxl.styles import PatternFill
import xlwings as xw
import zipfile
from io import StringIO
import os
from collections import defaultdict
from transfer_time import replace_all_spaces
from transfer_time import find_name_range
from transfer_time import range_adjustment
from itertools import chain

from openpyxl.worksheet.worksheet import Worksheet


# open file that will be used create billing docs.
def open_word_file():
    directory = filedialog.askopenfile()
    pass


# open Excel file that has the information about the extra charges.
def open_excel_file() -> str:
    return filedialog.askopenfilename(title='打刻表を選択してください。')


# create list or dict with all the extra charges for each children.
def count_charges():
    file_path = open_excel_file()
    book = openpyxl.load_workbook(file_path, keep_vba=False, data_only=True)
    charges = defaultdict(lambda : defaultdict(list))
    # iterate through the sheets
    for sheet_name in book.sheetnames[2:11]:
        sheet = book[sheet_name]
        row_ranges = find_name_range(sheet)
        # iterate through the two ranges (its like doing two chained iterations. but this way it's easier to calculate
        # the date row)
        for row_range in row_ranges:
            start = row_range[0]
            end = row_range[1]
            for row in sheet.iter_rows(min_row=start, max_row=end-1):
                name = row[2].value
                for i, cell in enumerate(row[5::4]):
                    price = cell.value
                    if price is not None and price >= 100:
                        date_row = start - 4
                        date_col = i*4 + 4 # (its 6 for the price but 2 less for the column that has the date.)
                        date = sheet.cell(row=date_row, column=date_col).value
                        print(sheet_name, name, price, date)



if __name__ == '__main__':
    count_charges()
