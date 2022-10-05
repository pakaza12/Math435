import xlrd

# Open the Workbook

def open_workbook(workbookName, sheetNum):
    workbook = xlrd.open_workbook(workbookName)
    a = workbook.sheet_by_index(sheetNum)

    worksheeta = []
    for i in range(0, a.nrows):
        worksheeta.append([])
        for j in range(0, a.ncols):
            worksheeta[i].append(a.cell_value(i, j))
    return worksheeta

def intTryParse(value):
    try:
        return int(value), True
    except ValueError:
        return value, False

def tryFilterColumn(columnNum, worksheet, filterVal):
    numRemoved = 0;
    for i in range(1, len(worksheet)):
        if float(worksheet[i-numRemoved][columnNum]) <= float(filterVal):
            worksheet.remove(worksheet[i-numRemoved])
            numRemoved += 1;
    return worksheet


"""
for i in range(0, data.nrows):
    for j in range(0, data.ncols):
        print(data.cell_value(i, j), end='\t')
    print('\n')
"""

excelOptions = ["MitC2006data.xlsx", "MitC2012data.xls", "MitC2022data - SalesPopulation.xlsx", "MitC2022data - VacantSales.xlsx"]
worksheetOptions = ["Linear Regression", "Multivariate Regression", "Clean Data"]
columnOptions = ["Remove Fields >=", "Remove Fields <=", "Remove fields =", "Remove Fields >", "Remove Fields <"]

options = [excelOptions,worksheetOptions,columnOptions]

worksheetHistory = []

while True:
    print("What xlsx would you like to work with?")
    for i in range(1, len(excelOptions)+1):
        print(str(i) + ".) " + excelOptions[i-1])

    excelSelection = input()

    if not intTryParse(excelSelection)[1] or int(excelSelection) <= 0 or int(excelSelection) > len(excelOptions):
        break

    worksheet = open_workbook(excelOptions[int(excelSelection)-1], 0)
    print(type(worksheet))

    worksheetHistory.append(worksheet)

    while True:

        print("\nWhat would you like to do with the worksheet?")
        for i in range(1, len(worksheetOptions)+1):
            print(str(i) + ".) " + worksheetOptions[i-1])

        worksheetSelection = input()

        if worksheetSelection != "1" and worksheetSelection != "2" and worksheetSelection != "3":
            break

        while True:
            print("\nWhich column would you like to clean?")
            for i in range(1, len(worksheet[0])):
                print(str(i) + ".) " + worksheet[0][i-1])

            cleanSelection = input()

            if not intTryParse(cleanSelection)[1] or int(cleanSelection) <= 0 or int(cleanSelection) > len(worksheet[0]):
                break

            colName = worksheet[0][int(cleanSelection)-1]
            print("\nColumn Selected:" + colName)


            print("\nHow would you like to clean this column?")
            for i in range(1, len(columnOptions)+1):
                print(str(i) + ".) " + columnOptions[i-1])

            cleanOptSelection = input()

            if not intTryParse(cleanOptSelection)[1] or int(cleanOptSelection) <= 0 or int(cleanOptSelection) > len(columnOptions):
                break


            print("\nEnter value to filter " + colName)

            filterSelection = input()

            a = tryFilterColumn(int(cleanSelection)-1, worksheet, filterSelection)
            print(a)


        #print("Enter what value you would like to filter by")
        #print(type(worksheet.cell_value(1, int(cleanSelection)-1)))
