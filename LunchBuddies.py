import sys
import os


def usage() -> None:
    print(r"""
    Improper command line usage!
   
    PROPER COMMAND LINE USAGE:
        py.exe .\LunchBuddies.py excel-file-path numLunchBuddyPairings

    EXAMPLE USAGE:
        py.exe .\LunchBuddies.py "LunchBuddiesInfo.xlsx" 3
        py.exe .\LunchBuddies.py "C:\Users\User.Name\Desktop\LunchBuddiesInfo.xlsx" 13

    BAD USAGE:
        py.exe .\LunchBuddies.py 
        py.exe .\LunchBuddies.py 3 
        py.exe .\LunchBuddies.py "LunchBuddiesInfo.xlsx" 
""")


def openExcelFile(excelFile):
    pass


if __name__ == "__main__":
    if len(sys.argv) != 3:
        usage()
        sys.exit(1)

    excelFile = sys.argv[1]
    numLunchBuddyPairings = sys.argv[2]

    if not os.path.isfile(excelFile):
        print(f"The file {excelFile} does not exist.")
        sys.exit(1)

    # dictionary of executive member names and general member names
    names = {}
    names = openExcelFile(excelFile)

    



