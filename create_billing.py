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
from openpyxl.styles import Color
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
def open_billing_file(option: int) -> str:
    if option == 1:
        title = '料金明細票を選択してください。'
    elif option == 2:
        title = '料金集計を選択してください。'
    else:
        title = 'incorrect option chosen'

    return filedialog.askopenfilename(title=title)



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
    color_map = {'ひよこ': 'FFFFFCE6', 'ひつじ': 'FFE6FFE6', 'うさぎ': 'FFE6FFFF', 'もも': 'FFFFC0CB',
                 'だいだい': 'FFFFA500', 'き': 'FFFFFF00', 'みどり': 'FF008000', 'あお': 'FF0000FF', 'ふじ': 'FF800080'}
    color = Color(rgb=color_map[class_name])
    sheet.sheet_properties.tabColor = color

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
    class_age_map = {'あお': '5', 'ふじ': '5', 'き': '4', 'みどり': '4', 'だいだい': '3', 'もも': '3',
                     'うさぎ': '2', 'ひつじ': '1', 'ひよこ': '0'}
    full_name = child_name.split('　')
    last = full_name[0]
    first = full_name[1]
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            if '%' in cell.value:
                cell.value = cell.value.replace('%', str(convert_reiwa(year, month)))
            if '#' in cell.value:
                cell.value = cell.value.replace('#', str(month))
            if '?' in cell.value:
                cell.value = cell.value.replace('?', class_age_map[class_name])
            if '@' in cell.value:
                cell.value = cell.value.replace('@', class_name)
            if '&' in cell.value:
                cell.value = cell.value.replace('&', last)
            if '$' in cell.value:
                cell.value = cell.value.replace('$', first)


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
    month = ye_mo_da[1]
    date = ye_mo_da[2]
    return f'{month}月{date}日'


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
def recalc_number(formula: str, num_rows_inserted: int, cell_range: bool) -> tuple[int, int, int]:
    """
    This function finds the specific number in the formula that need to be changed. Depending on whether the formula
    is a simple cell number or a formula that uses a range of cells it looks for the desired number to change as well
    as its location.

    :param formula: this is the formula that was found in the cell that need to be adjusted.
    :param num_rows_inserted: This basically is the number of days the child was charged extra.  This is because
    for every day the child was charged, there will be a new row inserted to record the charge.
    :param range: a bool argument that determines how we search for the number that need to be change. If True it
    searches the formula assuming that it uses as range of cells.(ex: sum(B23:B25))  If False it searched the formula assuming that it
    uses a single cell. (ex: =B23)
    :return: returns a tuple containing [new_num, start, end]. new_num is the adjusted number after recalculating with
    the numbers of rows inserted, start is the index of where the number started and end is the index for where the
    number ended. these will be used to splice together the new formula that will be inserted into as cell in the
    new sheet.
    """
    start = None
    end = None
    if cell_range is True:
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
def adjust_formulas(sheet: Worksheet, cells_to_be_adjusted: tuple[tuple[int, int, bool]], num_rows_inserted: int) -> None:
    """
    Because rows are being inserted, the range of the formulas that tally the total amounts need to adjust for that.
    Normally if we do this in Excel is automatically adjusts it. However, with openpyxl when a row is inserted the
    formulas are not adjusted automatically.  That's why, depending on the number of rows inserted we need to adjust
    the formulas present in the Excel book.

    :param sheet: new sheet that was created for each child that has extra charges.
    :param cells_to_be_adjusted: tuple of tuples containing the rows, columns, and type of formula of the cells that
    need their formulasv to be adjusted. ex:((3, 2, True), (10, 5, False)) Each tuple in the tuple represents a cell.
    The first number in the tuple is the row, the second number is the column of the cell that need to be changed and
    the last bool represents if the formula uses a range of cells or just a single cell. (For True it is a
    range =SUM(D3:D10), for False it is a single cell =D43.
    :param num_rows_inserted: This basically is the number of days the child was charged extra.  This is because
    for every day the child was charged, there will be a new row inserted to record the charge.
    :return:
    """
    for row, column, cell_range in cells_to_be_adjusted:
        formula = sheet.cell(row, column).value

        new_num, start, end = recalc_number(formula, num_rows_inserted, cell_range)
        sheet.cell(row, column).value = formula[:start] + str(new_num) + formula[end:]


# one of three documents that this automation creates.  This one creates the billing documents that will be
# given to the parents with all the individual charges organized by day.
def create_billing_sheets(charges: defaultdict, year: int, month: int) -> None:
    file_path = open_billing_file(1)
    book = openpyxl.load_workbook(file_path, keep_vba=False)
    sheet = book[book.sheetnames[0]]

    for class_name in charges:
        for kid_name in charges[class_name]:
            print(month, class_name, replace_all_spaces(kid_name))
            new_sheet_name = f'{month}月{class_name}{replace_all_spaces(kid_name)}'
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
                cells_to_be_adjusted =((16, 7, True), (30, 4, False))
                formula_1 = sheet.cell(16, 7).value
                formula_2 = sheet.cell(30, 4).value
                adjust_formulas(new_sheet, cells_to_be_adjusted, rows_inserted - 1)
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


def price_per_child_total(child_charges: defaultdict) -> int:
    total = 0
    for data in child_charges:
        total += data[0]
    return total


def insert_tally_data(new_sheet: Worksheet, row: int, class_name: str, kid_name: str, price: int) -> None:
    class_age_map = {'あお': '5', 'ふじ': '5', 'き': '4', 'みどり': '4', 'だいだい': '3', 'もも': '3',
                     'うさぎ': '2', 'ひつじ': '1', 'ひよこ': '0'}
    for cells in new_sheet.iter_rows(min_row=row, max_row=row):
        cells[0].value = f'{class_age_map[class_name]}歳児'
        cells[1].value = class_name
        cells[2].value = kid_name
        cells[3].value = price


def insert_formular_class_total(new_sheet: Worksheet, rows_inserted: int, first_row: int) -> int:
    last_row = 3 + rows_inserted
    new_sheet.cell(row=last_row, column=8).value = f'=SUM(D{first_row}:D{last_row})'
    return last_row + 1
# temporary place-holder for a function to create the second document I need.
def create_tally_sheet(charges: defaultdict, year: int, month: int) -> None:
    file_path = open_billing_file(2)
    book = openpyxl.load_workbook(file_path)
    sheet = book['base']
    new_sheet_name = f'{year}.{month}'
    new_sheet = book.create_sheet(new_sheet_name)
    copy_sheet(sheet, new_sheet)
    new_sheet.cell(row=2, column=1).value = f'{year}.{month}月'
    merge_cells(sheet, new_sheet)
    copy_dimensions(sheet, new_sheet)

    first = True
    count = 0
    rows_inserted = 0
    first_row = 3 # row number that will be used to set the range for the sum of each class.
    for class_name in charges:
        for kid_name in charges[class_name]:
            print(class_name, kid_name)
            price = price_per_child_total(charges[class_name][kid_name])
            if first is not True:
                new_sheet.insert_rows(count + 3)
                rows_inserted += 1
            copy_row_contents(new_sheet, 3, count + 3)
            insert_tally_data(new_sheet, count + 3, class_name, kid_name, price)
            first = False
            count += 1
        # this functions returns an object as well as mutates the new_sheet that is passed in.
        first_row = insert_formular_class_total(new_sheet, rows_inserted, first_row)
    cells_to_be_adjusted = ((4 + rows_inserted, 4, True), (4 + rows_inserted, 7, True), (4 + rows_inserted, 8, True))
    adjust_formulas(new_sheet, cells_to_be_adjusted, rows_inserted)

    book.save(new_file_path(file_path))


# main function to run all the processes I need.  Currently, this only create one file because I need to think about
# how I want to organize my code.  Later on I want to make it so my scripts create 3 different files, but some of these
# might need someone to manually place in data, so it doesn't make sense to automate it quite yet. I might actually need
# to make separate executables for each one so that someone can do the editing and a chose when to move onto the next
# automation portion of the work flow.
def main():
    charges = count_charges()
    year = find_year(charges)[0]
    month = find_year(charges)[1]
    #create_billing_sheets(charges, year, month)
    create_tally_sheet(charges, year, month)


# non important function to count how many children will be parched extra for the month.
def testtest(dic: dict) -> None:
    count = 0
    for clas in dic:
        for name in clas:
            count += 1
    print(count)


if __name__ == '__main__':
    #count_charges()
    main()
    #testtest(count_charges())
    #find_year(count_charges())
    #convert_reiwa(2024, 4)