# Transfer_time_automation
 transfers time stamps from one excel file to another


Version 1.7

transfers drop off and pick up times from one excel file to another that then calculates
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
 9 -- fixed the issue where the VBA code needed to be recalculated by opening the excel file in excel by triggering the
      recalculation within python.  Also made it so that the excel opening up is invisible to make it cleaner.
10 -- No longer need to physically choose the recalculated excel file during execusion, it is automaticallt passed into
      the function that fills in the cells with extra charges.
(new)
11 -- Made it so the reference files (excel documents downloaded from hug note that has the time stamps of arrival and
      departure times of all the kids) can be opened regardless if they are zipped or unzipped.
12 -- Fixed the issue where dep_check_time was being applied to all children. We only want to apply this to
      children that are 一号.
(new)
13 -- Iterate through the excel file to fill in cells that have both the arrival time and departure time blank with
      休み to indicate that the child did not come to school on that day. Also highlight with yellow on cells that have
      only arrival time or departure time missing but not both to indicate something went wrong or the parents forgot
      to record the time for arrival or departure.
(working on)
** -- Finish type hints and doc strings for all the functions.
** -- Fixed issue where 一号課外 time adjustments were being made every single week. its not every week that they have
      課外 classes, some are twice a month and some get canceled for one reason or another.
** -- Cleaned up code so that 0 index and 1 index difference between enumerate and the workbook are taken care of within
      their respective functions.
** -- Other various cleanups to make the code readable as well as organize things and changes to speed up things.

