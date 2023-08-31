# Transfer_time_automation
 transfers time stamps from one excel file to another


Version 1.4

First stable version that transfers drop off and pick up times from one excel file to another that then calculates
the appropriate amount of money to charge.

General features
1 -- This version transfers time from one excel file to another.
2 -- There is a window interface to choose the files needed in this script.
3 -- If there are missing files from the refrerence data (data downloaded from hugh note) it will notify you that not all
     files have been downloaded from hugnote, but will transfer data with what ever files are available.
4 -- If children on the hugnote files(reference files) cannot be found in the 預かり料金表 then it will notify you which
     children cannot be found along with the class name.
5 -- Converts the pick up time if a child is in 課外授業 and is 一号.  Children that are 一号 taking the after school class
     are exempt for charges resulting for being picked up late, up to a certain point.
6 -- created function to replace spaces between names including the japanese space aka IDEOGRAPHIC SPACE character or
     \u3000 in unicode escape character.  This was done to reduce replication to reduce effort when refactoring code.
(new)
7 -- Iterate through the excel file to find where we charged extra money and fill in those cells with a pink color,
     to make it easier to find where we charged extra.
8 -- fixed it so that the workbooks are properly closed at the end of the function to prevent any unwanted things
     from happnening with other functions down the line.


!!!WARNING!!!
issues to fix

1 -- this version has an issue where i need to reopen the excel file to color in the cells where charges were made for
     all the children.  The issue is that this charge is calculated internally by a custom VBA code within excel.  if you
     try to open the file through python the VBA code is not executed and thus the cells with the charges cannot be colored
     in.  To solve this problem you can open the _result file, enable macros, save and close the file before you chose the
     file to be opened for color in part of the code.  (that is coloring the cells that contian charges)
