import xlrd
import random
import statsmodels.api as sm
import seaborn as sns
import matplotlib.pyplot as plt
import numpy as np
import sklearn.linear_model as lm
import xlsxwriter
import string
import datetime

# Open the Workbook

def open_workbook(workbookName, sheetNum):
    if workbookName == "Other":
        print("What file would you like to open?")
        workbookName = input()

    workbook = xlrd.open_workbook(workbookName)
    a = workbook.sheet_by_index(sheetNum)

    worksheet = []
    for i in range(0, a.nrows):
        worksheet.append([])
        for j in range(0, a.ncols):
            if str(a.cell_value(0, j)).__contains__("Date") and i > 0:
                dateVal = a.cell_value(i, j)
                # print(dateVal)
                date = datetime.datetime(*xlrd.xldate_as_tuple(dateVal, workbook.datemode))
                # print(date)
                worksheet[i].append(date.date())
            else:
                worksheet[i].append(a.cell_value(i, j))

    return worksheet

def save_worksheet(worksheet):
    print("What would you like to name the file?")
    worksheetName = input()

    workbook = xlsxwriter.Workbook(worksheetName + '.xlsx')

    sheet = workbook.add_worksheet()

    uppercase_alphabets = list(string.ascii_uppercase)

    for i in range(0, len(worksheet)):
        for j in range(0, len(worksheet[i])):
            letter = ""
            if j > 25:
                letter = uppercase_alphabets[0] + uppercase_alphabets[j%26] + str(i+1)
            else:
                letter = uppercase_alphabets[j] + str(i+1)
            sheet.write(letter, worksheet[i][j])

    workbook.close()

def split_excel_files(worksheet):
    twentyPercentRows = int((len(worksheet)-1) * 0.2)

    copyWorksheet = []

    for i in range(0, len(worksheet)):
        copyWorksheet.append([])
        for j in range(0, len(worksheet[i])):
            copyWorksheet[i].append(worksheet[i][j])

    twentyPercent = [[]]
    for j in range(0, len(worksheet[0])):
            twentyPercent[0].append(worksheet[0][j])

    for i in range(1, twentyPercentRows+1):
        twentyPercent.append([])
        randRange = range(1, len(copyWorksheet)-1)
        randRow = random.choice(randRange)

        for j in range(0, len(worksheet[i])):
            twentyPercent[i].append(copyWorksheet[randRow][j])

        copyWorksheet.remove(copyWorksheet[randRow])

    print()
    print("Eighty Percent Data: ")
    save_worksheet(copyWorksheet)
    print()
    print("Twenty Percent Data: ")
    save_worksheet(twentyPercent)

    return

def intTryParse(value):
    try:
        return int(value), True
    except ValueError:
        return value, False
    except TypeError:
        return value, False

def floatTryParse(value):
    try:
        return float(value), True
    except ValueError:
        return value, False
    except TypeError:
        return value, False

def stringTryParse(value):
    try:
        return str(value), True
    except ValueError:
        return value, False
    except TypeError:
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
    if compareType == "replace":
        val2 = val2.split(",")[0]
    val1, val2, canCompare = convertSameTypes(val1, val2)
    if canCompare:
        if compareType == "contains":
            return str(val1).__contains__(str(val2))
        if compareType == "notContains":
            return not str(val1).__contains__(str(val2))
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
        if compareType == "replace":
            return val1 == val2
    return False

def tryFilterColumn(columnNum, worksheet, filterVal, compareType):

    numRemoved = 0;
    for i in range(1, len(worksheet)):
        if comparator(compareType, worksheet[i-numRemoved][columnNum], filterVal):
            if (compareType != "replace"):
                worksheet.remove(worksheet[i-numRemoved])
                numRemoved += 1
            else:
                # print("before: " + str(worksheet[i-numRemoved][columnNum]))
                worksheet[i-numRemoved][columnNum] = filterVal.split(",")[1]
                # print("after: " + str(worksheet[i-numRemoved][columnNum]))
        # else:
            # print(worksheet[i-numRemoved][columnNum])

    print("Number of Rows Removed: " + str(numRemoved))

    return worksheet

def clean_data(worksheet):
    while True:
        print("\nWhich column would you like to clean?")
        for i in range(1, len(worksheet[0])+1):
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

        if int(cleanOptSelection) == 6:
            worksheet = tryFilterColumn(int(cleanSelection)-1, worksheet, '', "equals")
            break

        print("\nEnter value to filter " + colName + ":")

        filterSelection = input()

        worksheet = tryFilterColumn(int(cleanSelection)-1, worksheet, filterSelection, compareType)
        worksheetHistory.append(worksheet)
        break
    return worksheet;

def linear_regression(worksheet):
    while True:
        print("\nSelect your first column (Independent Variable):")
        for i in range(1, len(worksheet[0])+1):
            print(str(i) + ".) " + worksheet[0][i-1])

        columnSelection1 = input()

        if not intTryParse(columnSelection1)[1] or int(columnSelection1) <= 0 or int(columnSelection1) > len(worksheet[0]):
            break

        colName1 = worksheet[0][int(columnSelection1)-1]

        print("\nSelect your second column (Dependent Variable):")
        for i in range(1, len(worksheet[0])+1):
            print(str(i) + ".) " + worksheet[0][i-1])

        columnSelection2 = input()

        if not intTryParse(columnSelection2)[1] or int(columnSelection2) <= 0 or int(columnSelection2) > len(worksheet[0]) or columnSelection1 == columnSelection2:
            break

        colName2 = worksheet[0][int(columnSelection2)-1]

        yData = []
        xData = []

        try:

            for i in range(1, len(worksheet)):
                # print("a: " + str(worksheet[i][int(columnSelection1)-1]))
                # print("b: " + str(worksheet[i][int(columnSelection2)-1]))
                yData.append(float(worksheet[i][int(columnSelection1)-1]))
                xData.append(float(worksheet[i][int(columnSelection2)-1]))

            x = np.array(xData).reshape((-1, 1))
            y = np.array(yData)

            reg = lm.LinearRegression()
            reg.fit(x, y)

            y_pred = reg.predict(x)

            # Coefficient of determination
            r_squared = reg.score(x, y)
            print("R-Squared: " + str(r_squared))

            # slope
            slope = reg.coef_
            print("Slope: " + str(slope))

            # intercept
            intercept = reg.intercept_
            print("Intercept: " + str(intercept))

            print("Equation: y=" + str(slope) + "x+" + str(intercept))

            sns.set_style('darkgrid')        # darkgrid, white grid, dark, white and ticks
            plt.rc('axes', titlesize=23)     # fontsize of the axes title
            plt.rc('axes', labelsize=20)     # fontsize of the x and y labels
            plt.rc('xtick', labelsize=16)    # fontsize of the tick labels
            plt.rc('ytick', labelsize=16)    # fontsize of the tick labels
            plt.rc('legend', fontsize=16)    # legend fontsize
            plt.rc('font', size=16)          # controls default text sizes


            sns.scatterplot(x=xData, y=yData, color='blue')
            plt.xlabel(colName2)
            plt.ylabel(colName1)
            sns.lineplot(x=xData, y=y_pred, color='red')
            plt.xlim(0)
            plt.ylim(0)
            plt.show()

            break

        except:
            print("Could not properly compare columns")
            break

    return

def multiple_linear_regression(worksheet):
    while True:
        print("\nSelect your independent variable(s). You can enter multiple by separating them with a comma. Ex: 1,2")
        for i in range(1, len(worksheet[0])+1):
            print(str(i) + ".) " + worksheet[0][i-1])

        columnSelection1 = input()

        if len(columnSelection1.split(",")):
            columnSelection1 = columnSelection1.split(",")
        else:
            columnSelection1 = [columnSelection1]

        # print(columnSelection1)

        isValidInput = True
        for i in range(0, len(columnSelection1)):
            if not intTryParse(columnSelection1[i])[1] or int(columnSelection1[i]) <= 0 or int(columnSelection1[i]) > len(worksheet[0]):
                isValidInput = False

        if not isValidInput:
            break

        colNames1 = []

        for i in range(0, len(columnSelection1)):
            colNames1.append(worksheet[0][int(columnSelection1[i])-1])

        print("\nSelect your second column (Dependent Variable):")
        for i in range(1, len(worksheet[0])+1):
            print(str(i) + ".) " + worksheet[0][i-1])

        columnSelection2 = input()

        if not intTryParse(columnSelection2)[1] or int(columnSelection2) <= 0 or int(columnSelection2) > len(worksheet[0]) or columnSelection2 in columnSelection1:
            break

        colName2 = worksheet[0][int(columnSelection2)-1]

        xData = []
        yData = []
        xTest = []
        yTest = []
        twentyPercentRows = int((len(worksheet)-1) * 0.2)

        try:
            for i in range(1, len(worksheet)):
                xData.append([])
                for j in range(0, len(columnSelection1)):
                    xData[i-1].append(float(worksheet[i][int(columnSelection1[j])-1]))

            for i in range(1, len(worksheet)):
                yData.append(float(worksheet[i][int(columnSelection2)-1]))


            for i in range(0, twentyPercentRows):
                xTest.append([])
                randRange = range(0, len(xData)-1)
                randRow = random.choice(randRange)

                yTest.append(yData[randRow])
                yData.remove(yData[randRow])

                for j in range(0, len(columnSelection1)):
                    xTest[i].append(xData[randRow][j])
                xData.remove(xData[randRow])

            print()
            print("xTest Length: ")
            print(len(xTest))
            print("yTest Length: ")
            print(len(yTest))

            print("xData Length: ")
            print(len(xData))
            print("yData Length: ")
            print(len(yData))
            print()

            reg = lm.LinearRegression()
            reg.fit(xData, yData)
            slope = reg.coef_

            print("slope: ")
            print(slope)
            intercept = reg.intercept_
            print("intercept: ")
            print(intercept)

            # y_pred = reg.predict(xTest)

            # print("y_pred: ")
            # print(y_pred)

            # Coefficient of determination
            r_squared = reg.score(xTest, yTest)
            print("R-Squared: " + str(r_squared))

            break

        except:
            print("Could not properly compare columns")
            break

    return

excelOptions = ["MitC2006data.xlsx", "MitC2012data.xls", "MitC2022data - SalesPopulation.xlsx", "MitC2022data - VacantSales.xlsx", "Other"]
worksheetOptions = ["Linear Regression", "Multivariate Regression", "Clean Data", "Save as xlsx", "Split Data 80/20"]
columnOptions = ["Remove Fields >=", "Remove Fields <=", "Remove fields =", "Remove Fields >", "Remove Fields <", "Remove Empty", "Contains", "Doesn't contain", "Replace _ with _ (Ex: 1,2 or '',test"]
columnDict = {1:"greaterThanEquals", 2:"lessThanEquals", 3:"equals", 4:"greaterThan", 5:"lessThan", 7:"contains", 8:"notContains", 9:"replace"}

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

        if worksheetSelection != "1" and worksheetSelection != "2" and worksheetSelection != "3" and worksheetSelection != "4" and worksheetSelection != "5":
            break

        if worksheetSelection == "3":
            worksheet = clean_data(worksheet)

        if worksheetSelection == "1":
            linear_regression(worksheet)

        if worksheetSelection == "2":
            multiple_linear_regression(worksheet)

        if worksheetSelection == "4":
            save_worksheet(worksheet)

        if worksheetSelection == "5":
            split_excel_files(worksheet)



        #print("Enter what value you would like to filter by")
        #print(type(worksheet.cell_value(1, int(cleanSelection)-1)))
