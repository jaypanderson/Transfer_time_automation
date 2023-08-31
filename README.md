# Transfer_time_automation
 transfers time stamps from one excel file to another


Version 1.2

First stable version that transfers drop off and pick up times from one excel file to another that then calculates
the appropriate amount of money to charge.

General features
1 -- This version transfers time from one excel file to another.
2 -- There is a window interface to choose the files needed in this script.
3 -- If there are missing files from the refrerence data (data downloaded from hugh note) it will notify you that not all
     files have been downloaded from hugnote, but will transfer data with what ever files are available.
4 -- If children on the hugnote files(reference files) cannot be found in the 預かり料金表 then it will notify you which
     children cannot be found along with the class name.
(new)
5 -- Converts the pick up time if a child is in 課外授業 and is 一号.  Children that are 一号 taking the after school class
     are exempt for charges resulting for being picked up late, up to a certain point.
