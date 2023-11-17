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
from collections import Counter
from itertools import zip_longest
from openpyxl.utils.cell import range_boundaries


from openpyxl.worksheet.worksheet import Worksheet


# open file that will be used create billing docs.
def open_billing_file() -> str:
    return filedialog.askopenfilename(title='料金明細票を選択してください。')



# open Excel file that has the information about the extra charges.
def open_excel_file() -> str:
    return filedialog.askopenfilename(title='打刻表を選択してください。')


# add result to the end of the file name
def new_file_path(path: str) -> str:
    idx = path.find('.')
    ans = path[:idx] +'result' + path[idx:]
    return ans


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
    print(charges)
    return charges


def set_color(sheet: Worksheet, class_name: str) -> None:
    color_map = {'ひよこ': '', 'ひつじ': '', 'うさぎ': '', 'もも': '',
                 'だいだい': '', 'き': '', 'みどり': '', 'あお': '', 'ふじ': ''}


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


# replicate the cell merges from base template.
def merge_cells(sheet: Worksheet, new_sheet: Worksheet) -> None:
    for merged_cell_range in sheet.merged_cells.ranges:
        new_sheet.merge_cells(str(merged_cell_range))


def copy_print_area(sheet: Worksheet, new_sheet: Worksheet) -> None:
    if sheet.print_area:
        new_sheet.print_area = sheet.print_area


def copy_dimensions(sheet: Worksheet, new_sheet: Worksheet) -> None:
    for row, col in zip_longest(sheet.row_dimensions, sheet.column_dimensions):
        # if row is not None:
        #     new_sheet.row_dimensions[row].height = sheet.row_dimensions[row].height
        if col is not None:
            new_sheet.column_dimensions[col].width = sheet.column_dimensions[col].width


# spliting child name into first and last name
def separate_names(name: str) -> list[str, str]:
    return name.split('　')


def insert_name_date(sheet: Worksheet, year: int, month: int, class_name: str, child_name: str) -> None:
    class_age_map = {'あお': 5, 'ふじ': 5, 'き': 4, 'みどり': 4, 'だいだい': 3, 'もも': 3, 'うさぎ': 2, 'ひつじ': 1, 'ひよこ': 0}
    full_name = child_name.split('　')
    last = full_name[0]
    first = full_name[1]
    for row in sheet.iter_rows():
        for cell in row:
            val = cell.value
            if '%' in val:
                cell.value = val.replace('%', year)
            if '#' in val:
                cell.value = val.replace('#', month)
            if '?' in val:
                cell.value = val.replace('?', class_age_map[class_name])
            if '@' in val:
                cell.value = val.replace('?', class_name)
            if '&' in val:
                cell.value = val.replace('&', separate_names(child_name)[0])
            if '$' in val:
                cell.value = val.replace('&', separate_names(child_name)[1])


def find_max(counts: Counter) -> int:
    highest = 0
    ans = None
    for i in counts:
        if counts[i] > highest:
            highest = counts[i]
            ans = i
    return ans


def find_year(charges: dict) -> tuple[int, int]:
    years = []
    months = []
    for i in charges:
        for j in charges[i]:
            for data in charges[i][j]:
                year = data[3].split('-')[0]
                years.append(year)
                month = data[3].split('-')[1]
                months.append(month)

    year_ans = find_max(Counter(years))
    month_ans = find_max(Counter(months))
    print(year_ans, month_ans)
    return int(year_ans), int(month_ans)


def convert_reiwa(year: int, month: int) -> int:
    reiwa = year - 2018
    if month in [1, 2, 3]:
        reiwa -= 1
    print(reiwa)
    return reiwa


def copy_row_contents(sheet: Worksheet, row_num: int, new_row_num: int) -> None:
    for row, new_row in zip(sheet.iter_rows(min_row=row_num, max_row=row_num),
                            sheet.iter_rows(min_row=new_row_num, max_row=new_row_num)):
        for cell, new_cell in zip(row, new_row):
            new_cell.font = copy(cell.font)
            new_cell.border = copy(cell.border)
            new_cell.fill = copy(cell.fill)
            new_cell.number_format = copy(cell.number_format)
            new_cell.protection = copy(cell.protection)
            new_cell.alignment = copy(cell.alignment)

def convert_date(date: str) -> str:
    ye_mo_da = date.split('-')
    mo_da_ye = [ye_mo_da[1], ye_mo_da[2], ye_mo_da[0]]
    return '/'.join(mo_da_ye)


def format_time(time: int) -> str:
    chars = list(str(time))
    if 3 > len(chars) > 4:
        return 'Time error'
    chars.insert(-2, ':')
    return ''.join(chars)


def insert_data(sheet: Worksheet, row: int, month: int, price: int, arrival: int, departure: int, date: str) -> None:
    for cells in sheet.iter_rows(min_row=row, max_row=row):
        cells[1].value = f'{month}月分預かり保育料金'
        cells[3].value = convert_date(date)
        cells[4].value = format_time(arrival)
        cells[5].value = format_time(departure)
        cells[6].value = price


def merge_specific_cells(sheet, new_row_num, start_col, end_col):
    merge_range = f'{start_col}{new_row_num}:{end_col}{new_row_num}'
    sheet.merge_cells(merge_range)


def adjust_merged_cells(sheet: Worksheet, loc_row_inserted, num_rows_inserted):
    new_merged_ranges = []
    for merged_range in sheet.merged_cells.ranges:
        min_col, min_row, max_col, max_row = range_boundaries(str(merged_range))

        if min_row >= loc_row_inserted:
            min_row += num_rows_inserted
            max_row += num_rows_inserted
        new_range = f'{openpyxl.utils.get_column_letter(min_col)}{min_row}:{openpyxl.utils.get_column_letter(max_col)}{max_row}'
        new_merged_ranges.append(new_range)

    sheet.merged_cells.ranges = []

    for new_range in new_merged_ranges:
        sheet.merge_cells(new_range)


# find the location of the number that needs to be recalculated and add the number of rows based on
# how many rows were inserted.
def recalc_number(formula: str, num_rows_inserted: int, range: bool) -> tuple[int, int, int]:
    start = None
    end = None
    if range is True:
        for i, char in enumerate(formula):
            if char == ':':
                start = i + 2  # it's two, to account for ':' and the column letter.
            if char == ')':
                end = i
    else:
        for i, char in enumerate(formula):
            if char.isalpha():
                start = i + 1
        end = len(formula)

    num = formula[start:end]
    new_num = int(num) + num_rows_inserted
    return new_num, start, end


# apply the new values to the formulas based on how many rows were inserted.
def adjust_formulas(sheet: Worksheet, num_rows_inserted: int) -> None:
    formula_1 = sheet.cell(16, 7).value
    formula_2 = sheet.cell(30, 4).value

    new_num, start, end = recalc_number(formula_1, num_rows_inserted, True)
    sheet.cell(16, 7).value = formula_1[:start] + str(new_num) + formula_1[end:]

    new_num, start, end = recalc_number(formula_2, num_rows_inserted, False)
    sheet.cell(30, 4).value = formula_2[:start] + str(new_num) + formula_2[end:]







def create_billing():
    file_path = open_billing_file()
    book = openpyxl.load_workbook(file_path, keep_vba=False)
    sheet = book[book.sheetnames[0]]

    charges = count_charges()
    year = find_year(charges)[0]
    month = find_year(charges)[1]
    for class_name in charges:
        for kid_name in charges[class_name]:
            print(month, class_name, replace_all_spaces(kid_name))
            new_sheet_name = f'{month}{class_name}{replace_all_spaces(kid_name)}'
            new_sheet = book.create_sheet(new_sheet_name)
            set_color(new_sheet, class_name)
            copy_sheet(sheet, new_sheet)
            merge_cells(sheet, new_sheet)
            copy_print_area(sheet, new_sheet)
            copy_dimensions(sheet, new_sheet)
            insert_name_date(new_sheet, year, month, class_name, kid_name)

            rows_inserted = len(charges[class_name][kid_name])
            if rows_inserted > 1:
                adjust_merged_cells(new_sheet, 15, rows_inserted - 1)
                adjust_formulas(new_sheet, rows_inserted - 1)
            for i, data in enumerate(charges[class_name][kid_name]):
                row_num = 14
                first_insertion_location = 15
                new_row_num = 14 + i
                if i != 0:
                    new_sheet.insert_rows(row_num + 1 + i)
                copy_row_contents(new_sheet, row_num, row_num + i)
                merge_specific_cells(new_sheet, row_num + i, 'B', 'C')
                insert_data(new_sheet, new_row_num, month, data[0], data[1], data[2], data[3])



    book.save(new_file_path(file_path))


def testtest(dic):
    count = 0
    for clas in dic:
        for name in clas:
            count += 1
    #print(count)


if __name__ == '__main__':
    #count_charges()
    create_billing()
    #testtest(count_charges())
    #find_year(count_charges())
    #convert_reiwa(2024, 4)