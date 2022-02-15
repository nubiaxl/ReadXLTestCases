# Reading an excel file using Python
import xlrd

# Give the location of the file
"""""
mypath = "/Users/nubilalevon/My-Documents/AllTestCases/"
filename = "GuruTest001v1.0.xlsx"
loc = mypath + filename
"""""

#User gives Column C and Row 1
def getTestCaseDataLocations():
    #call should be object.cellID
    cellID = parseLocations("Enter location for TestCaseID")
    cellTCDescription = parseLocations("Enter location for TestCase Description")


def parseLocations(prompt):
    try:
        column = input(prompt + " column")
        row  = input (prompt + " row")

        print("Column Number is: ", column)
        print("Second Number is: ", row )
        print()

        if len(column) > 1 or len(row) > 2:
            parseLocations(prompt)
        else:
            numColumn = ord(column) - ord('A')
            numRow = int(row)
            print ("Column: " + str(numColumn))
            return numRow, numColumn
    except IOError:
        print("IOError at parseLocations")

def getTestSuiteFiles(directory, regexpression, ext):
    import glob

    ## directory = 'c:\\projects\\hc2\\'
    ## regexpression of form "**/*"
    ## ext of form .xlsx .xls etc

    path = directory + regexpression + ext
    filterFilePathString(path)
    files = [f for f in glob.glob(path, recursive=True)]
    return files

    # for f in files:
     #   print(f)


def filterFilePathString(path):
# check to see that file exists
    path.replace("\\", "\\\\")
    print(path)

"""
cell = parseLocations("Enter TestCase ID Location: ")
print( cell[0] )
print ( cell[1] )


    descriptionLocation = input ("Enter location of Test Case Description")
    versionLocation = input ("Enter location of the Version")
    createdByLocation = input ("Enter location of the Author information")
    scenarioLocation = input ("Enter location of Scenario Description")



    stepDetailsRange = input ("Enter range of Step Details in form A1:A3")
    testDataRange = input ("Enter range of Test Data Values")

"""

def getCellValue(prompt, location):

    #cell needs to be global, part of object
    cell = parseLocations(prompt)
    # To open Workbook
    wb = xlrd.open_workbook(location)
    sheet = wb.sheet_by_index(0)

    # Get data from the excel cell
    sheetData =  sheet.cell_value(cell[0]-1, cell[1])
    return sheetData


directory = input("Please enter directory location of the test suite: ")
regex = input("Please enter your test filename regular expression form BUv1*")
extension = input("Please enter your excel filename extension (.xlsx): ")

loclist = getTestSuiteFiles(directory, regex, extension)
print(loclist)

#What if loc is a list of files?
for loc in loclist:
    testCaseID = getCellValue("Enter the location for TestCase ID: ", loc)
    print("TestCaseID= " + testCaseID)