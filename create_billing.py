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
from copy import copy

from openpyxl.worksheet.worksheet import Worksheet


# open file that will be used create billing docs.
def open_billing_file() -> str:
    return filedialog.askopenfilename(title='料金明細票を選択してください。')



# open Excel file that has the information about the extra charges.
def open_excel_file() -> str:
    return filedialog.askopenfilename(title='打刻表を選択してください。')


# create list or dict with all the extra charges for each child.
def count_charges():
    file_path = open_excel_file()
    book = openpyxl.load_workbook(file_path, keep_vba=False, data_only=True)
    charges = defaultdict(lambda : defaultdict(list))
    # iterate through the sheets
    for sheet_name in book.sheetnames[2:11]:
        sheet_name = replace_all_spaces(sheet_name) # replace spaces
        sheet = book[sheet_name]
        row_ranges = find_name_range(sheet)
        # iterate through the two ranges (its like doing two chained iterations. but this way it's easier to calculate
        # the date row)
        for row_range in row_ranges:
            start = row_range[0]
            end = row_range[1]
            for idx, row in enumerate(sheet.iter_rows(min_row=start, max_row=end-1)):
                name = row[2].value
                for i, cell in enumerate(row[5::4]):
                    price = cell.value
                    if price is not None and price >= 100:
                        date_row = start - 4
                        date_col = i*4 + 4 # (its 6 for the price but 2 less for the column that has the date.)
                        date = sheet.cell(row=date_row, column=date_col).value
                        date = str(date)[0:10]
                        arriv_row = start + idx
                        arriv_col = i * 4 + 4
                        dept_row = start + idx
                        dept_col = i * 4 + 5
                        arrival = sheet.cell(row=arriv_row, column=arriv_col).value
                        departure = sheet.cell(row=dept_row, column=dept_col).value
                        charges[sheet_name][name].append((price, arrival, departure, date))
                        #print(sheet_name, name, price, date)
    return charges

def copy_sheet(sheet, new_sheet) -> None:
    for row in sheet:
        for cell in row:
            new_cell = new_sheet.cell(row=cell.row, column=cell.column, value=cell.value)

            if cell.has_style:
                new_cell.font = copy(cell.font)
                new_cell.border = copy(cell.border)
                new_cell.fill = copy(cell.fill)
                new_cell.number_format = copy(cell.number_format)
                new_cell.protection = copy(cell.protection)
                new_cell.alignment = copy(cell.alignment)
                new_cell.comment = copy(cell.comment)



def create_billing():
    file_path = open_billing_file()
    book = openpyxl.load_workbook(file_path, keep_vba=True)
    source = book['base']

    charges = count_charges()
    temp = charges['あお']['宮西　つぐみ']
    month = '10月'
    class_name = 'あお'
    kid_name = '宮西　つぐみ'
    for _ in range(1):
        new_sheet = book.create_sheet(f'{month}{class_name}{replace_all_spaces(kid_name)}')
        copy_sheet(source, new_sheet)

    book.save(file_path + 'result')




if __name__ == '__main__':
    #count_charges()
    create_billing()
