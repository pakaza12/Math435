import xlrd

# Open the Workbook

def open_workbook(workbookName, sheetNum):
    workbook = xlrd.open_workbook(workbookName)
    a = workbook.sheet_by_index(sheetNum)

    worksheet = []
    for i in range(0, a.nrows):
        worksheet.append([])
        for j in range(0, a.ncols):
            worksheet[i].append(a.cell_value(i, j))
    return worksheet

def intTryParse(value):
    try:
        return int(value), True
    except ValueError:
        return value, False

def floatTryParse(value):
    try:
        return float(value), True
    except ValueError:
        return value, False

def stringTryParse(value):
    try:
        return str(value), True
    except ValueError:
        return value, False

def convertSameTypes(val1, val2):
    if floatTryParse(val1)[1] and floatTryParse(val2)[1]:
        return float(val1), float(val2), True
    if intTryParse(val1)[1] and intTryParse(val2)[1]:
        return int(val1), int(val2), True
    if stringTryParse(val1)[1] and stringTryParse(val2)[1]:
        return str(val1), str(val2), True
    return val1, val2, False

def comparator(compareType, val1, val2):
    val1, val2, canCompare = convertSameTypes(val1, val2)
    if canCompare:
        if compareType == "lessThan":
            return val1 < val2
        if compareType == "lessThanEquals":
            return val1 <= val2
        if compareType == "greaterThan":
            return val1 > val2
        if compareType == "greaterThanEquals":
            return val1 >= val2
        if compareType == "equals":
            return val1 == val2
    return False

def tryFilterColumn(columnNum, worksheet, filterVal, compareType):
    numRemoved = 0;
    for i in range(1, len(worksheet)):
        if comparator(compareType, worksheet[i-numRemoved][columnNum], filterVal):
            worksheet.remove(worksheet[i-numRemoved])
            numRemoved += 1;
        else:
            print(worksheet[i-numRemoved][columnNum])
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
columnDict = {1:"greaterThanEquals", 2:"lessThanEquals", 3:"equals", 4:"greaterThan", 5:"lessThan"}

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

            compareType = columnDict.get(int(cleanOptSelection))
            print("\nEnter value to filter " + colName + ":")

            filterSelection = input()

            worksheet = tryFilterColumn(int(cleanSelection)-1, worksheet, filterSelection, compareType)
            worksheetHistory.append(worksheet);



        #print("Enter what value you would like to filter by")
        #print(type(worksheet.cell_value(1, int(cleanSelection)-1)))
