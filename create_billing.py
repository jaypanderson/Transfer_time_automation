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

from openpyxl.worksheet.worksheet import Worksheet


# open file that will be used create billing docs.
def open_word_file():
    directory = filedialog.askopenfile()
    pass


# open Excel file that has the information about the extra charges.
def open_excel_file() -> str:
    return filedialog.askopenfile(title='打刻表を選択してください。')


# create list or dict with all the extra charges for each children.
def count_charges():
    file_path = open_excel_file()
    book = openpyxl.workbook(file_path, data_only=True)
    charges = defaultdict(lambda : defaultdict(list))
    for sheet in book.sheetnames[2:11]:
