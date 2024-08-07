# LunchBuddies

Purpose: The goal of this script is to take in an excel workbook that contains a sheet with 2 lists of names, one list for executive board members and the other 
of general member names. This script will generate x number of additional sheets in the workbook where an executive board name is paired with a general member name, and the priority is to pair executive board members with general members - whichever group has more people will just be paired internally within that group afterward. For the x number of sheets generated, each pairing is unique. That is, a pairing will never be repeated across these x sheets. Each sheet represents one week of a lunch buddies.

## Steps to run: ##
1. Download Python and openpyxl libraries (see Requirements)
2. Prepare excel file as described below.
3. Close out of the excel file - it will not be able to create new sheets if excel file is opened.
4. Run the script as shown below.

## Requirements ##
- Python 3.10.2 
- openpyxl : Once Python is downloaded, run `py -m pip install openpyxl==3.1.5`

## Excel File Setup ##
The excel file should have two columns. One to denote the Executive Members list and the second column to denote the General Members list. Cells A1 and B1 should have descriptive column names such as "Executive Members" and "General Members". The names should populate the rows underneath both columns. If there are an odd number of total members (executive members + general members), there will be one unpaired person each week. Ensure there is only this one sheet in this workbook. 

## How to use SRDRscript.py

PROPER COMMAND LINE USAGE:
- `py.exe .\LunchBuddies.py excel-file-path numWeeks`

EXAMPLE USAGE:
- `py.exe .\LunchBuddies.py "LunchBuddiesInfo.xlsx" 3`  
- `py.exe .\LunchBuddies.py "C:\Users\User.Name\Desktop\LunchBuddiesInfo.xlsx" 13`
    
BAD USAGE:
- Missing excel file and number of weeks: `py.exe .\LunchBuddies.py `  
- Incorrect order of arguments: `py.exe .\LunchBuddies.py 13 "LunchBuddiesInfo.xlsx" `  
- Missing excel file: `py.exe .\LunchBuddies.py 3`  
