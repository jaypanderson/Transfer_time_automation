"""
version 1.1

First stable version that transfers drop off and pick up times from one excel file to another that then calculates
the appropriate amount of money to charge.

General features
1 -- This version transfers time from one excel file to another.
2 -- There is a window interface to choose the files needed in this script.
3 -- If there are missing files from the refrerence data (data downloaded from hugh note) it will notify you that not all
     files have been downloaded from hugnote, but will transfer data with what ever files are available.
4 -- If children on the hugnote files(reference files) cannot be found in the 預かり料金表 then it will notify you which
     children cannot be found along with the class name.
"""

import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from datetime import datetime
import openpyxl
import os



def find_date(tab, date):
    for i, row in enumerate(tab.iter_rows()):
        for idx, cell in enumerate(row):
            if cell.value == date:
                return [i, idx]
    return None



def find_name(tab1, name, date_row):
    ans = []
    for i, row in enumerate(tab1.iter_rows()):
        if type(row[2].value) == str:
            cell_name = row[2].value.replace('　', '') # here i am replacing a japanese space with an empty string. The japanese space is different from the us space.
            cell_name = cell_name.replace(' ', '') # also replacing regular spaces.
            #print(cell_name, name)
            if cell_name == name:
                ans.append(i)
    if date_row < 10:
        return ans[:1] # returning just one so make sure we only place the time correctly for the corresponding date
    else:
        return ans[-1:] # also using [:1] and [-1:] so an error is not raised when the list is empty



def arr_check_time(time:str) -> int:
    time = int(time)
    if  time < 730:
        time = 730
    return time

def dep_check_time(time:str) -> int:
    time = int(time)
    if 1131 <= time <= 1139:
        time = 1130
    if 1431 <= time <= 1439:
        time = 1430
    return time



def update_excel_data(input_file, reference_files, output_file):


    # Read the input Excel file with openpyxl
    output_data = openpyxl.load_workbook(input_file, data_only=False, keep_vba=True)
    input_data = openpyxl.load_workbook(input_file, data_only=True)

    # Read the reference CSV files
    reference_data = {}
    for key, val in reference_files.items():
        #print(key)
        reference_data[key] = pd.read_csv(val, parse_dates=['日付'])

    missing_children = set()
    # Iterate over the input data tabs
    for sheet_name in input_data.sheetnames[2:]:

        # access the sheet we are currently working on
        cur_sheet = input_data[sheet_name]
        out_sheet = output_data[sheet_name]

        # erase any possible spaces in the sheetname
        new_sheet_name = sheet_name.replace(' ', '')  # replacing normal english space
        new_sheet_name = new_sheet_name.replace('　', '')  # replacing japanese space

        # check to see if tab name exists in reference data
        # if there is no match it is possible the user did not download all the files
        # from hugnote and needs to make sure all calsses are selected.
        if new_sheet_name not in reference_data:
            messagebox.showinfo('全てのクラスがダウンロードされてません。', '{}組がダウンロードされてません'.format(new_sheet_name))
            continue
        # Read the reference data for the current tab
        ref_data = reference_data[new_sheet_name]

        #print(cur_sheet.max_row)
        # Iterate through the refference data
        for i, row in ref_data.iterrows():
            date = row['日付']
            child_name = row['こども氏名'].replace(' ', '')
            child_name = child_name.replace('　', '') # also replace the Japanese space.
            arrive_time = row['出席時刻']
            departure_time = row['帰宅時刻']

            # remove : from time stamp and skip procedure if it is a nan value.
            if isinstance(arrive_time, str):
                arrive_time = arrive_time.replace(':', '')

            if isinstance(departure_time, str):
                departure_time = departure_time.replace(':', '')

            #find the corresponding date(cell row and col) date_coor[0] is the row and date_coor[1] is the column
            date_coor = find_date(cur_sheet, date)
            #print(i, date_coor, date)

            # check to see if date_coor is empty or is None, if so skip the date. because it can cause errors in the fillowing operations.
            if date_coor == None:
                continue

            # find the corresponding name. This only gives the row because we will use the column numbers from date_coor
            name_coor = find_name(cur_sheet, child_name, date_coor[0])
            if not name_coor:
                missing_children.add('{}組の{}'.format(new_sheet_name, child_name))
                print(i, date_coor, date)
                print(i, name_coor, child_name)

            # check to see if name_coor is an empty list. if so continue to next entry.
            if len(name_coor) == 0:
                continue

            # Write data into cells.
            if isinstance(arrive_time, str) :
                adj_arrive_time = arr_check_time(arrive_time)
                out_sheet.cell(name_coor[0] + 1, date_coor[1] + 1).value = adj_arrive_time # Add one to adjust for the 0 index created with the enumreate() function
            if isinstance(departure_time, str):
                adj_departure_time = dep_check_time(departure_time)
                out_sheet.cell(name_coor[0] + 1, date_coor[1] + 2).value = adj_departure_time # Add one to adjust for the 0 index created with the enumreate() function
    messagebox.showinfo('以下の子供が見つかりませんでした。', "ハグノートと預かり料金ファイルの子供の漢字が異なってる可能性があります。\n預かり料金ファイルの子供の名前を修正してください。:\n\n" + "\n".join(missing_children))





    output_data.save(output_file)


# create file paths by asking the user.

# Create the Tkinter root window
root = tk.Tk()
root.withdraw()  # Hide the root window


# prompt user for input file
input_file = filedialog.askopenfilename(title = '預かり料金表を選択してください。')
directory_path = filedialog.askdirectory(title = 'ダウンロードした打刻表のフォルダを選択してください。')
files = os.listdir(directory_path)

# Generate output file name
output_file = os.path.splitext(input_file)[0] + "_result.xlsm"

# create dictionary to store path names for reference files
class_names = ['ひよこ', 'ひつじ', 'うさぎ', 'だいだい',
              'もも', 'みどり', 'き', 'あお', 'ふじ']
reference_files = {}
for class_name in class_names:
    for file_name in files:
        file_path = os.path.join(directory_path, file_name) # create the new file path
        if os.path.isfile(file_path) and class_name in file_path:
            reference_files[class_name] = file_path



update_excel_data(input_file, reference_files, output_file)

