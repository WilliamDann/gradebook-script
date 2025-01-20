# import pandas lib as pd
import pandas as pd
from   sys    import argv
from export   import export

## TODO if args grows more complex, add a data structure to contain them
# returns (fileName, sheetName) for now, see above
def readCmdArgs():
    fileName  = None
    sheetName = None

    if len(argv) >= 2:
        fileName  = argv[1]
    if len(argv) >= 3:
        sheetName = argv[2]

    return (fileName, sheetName)

# Get a file name from the user, if missing
def fileNamePrompt():
    print("Please enter the path to your locally saved gradebook excel (.xlsx) file:")
    fileName = input("> ")

    try:
        pd.read_excel(fileName)
    except Exception as e:
        print("Invalid file:")
        print(e)
        return None

    return fileName

# get which sheet to use from the user
def sheetNamePrompt(dataframe):
    print("This gradebook has multiple classes, please enter the number for which sheet to export:")
    
    print()
    i = 0
    for sheet in dataframe.sheet_names:
        print(f"{i} - {sheet}")
        i += 1
    
    try:
        num = int(input("> "))
        return dataframe.sheet_names[num]
    except ValueError as e:
        print("Please only input a number")
        return None
    except IndexError as e:
        print("Please only input a number displayed ")
        return None

fileName, sheetName = readCmdArgs()

# Get a file name from the user if not found
while fileName is None:
    fileName = fileNamePrompt()

xlsx = pd.ExcelFile(fileName)

# Get a sheet name from the user if not found
while sheetName is None:
    sheetName = sheetNamePrompt(xlsx)

# send correct sheet to export function
export(xlsx.parse(sheetName), xlsx.parse("Grading"))