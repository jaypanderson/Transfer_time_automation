# Transfer_time_automation
 transfers time stamps from one excel file to another


Version 1.5

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
 7 -- Iterate through the excel file to find where we charged extra money and fill in those cells with a pink color,
      to make it easier to find where we charged extra.
 8 -- fixed it so that the workbooks are properly closed at the end of the function to prevent any unwanted things
      from happnening with other functions down the line.
(new)
 9 -- fixed the issue where the VBA code needed to be recalculated by opening the excel file in excel by triggering the
      recalculation within python.  Also made it so that the excel opening up is invisible to make it cleaner.
10 -- No longer need to physically choose the recalculated excel file during execution, it is automatically passed into
      the function that fills in the cells with extra charges.