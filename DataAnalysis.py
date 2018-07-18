__author__ = 'Timmy Desmond'

import csv

import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell

# Define excel color codes
red = '#FFC7CE'
yellow = '#FFF271'
orange = '#ff8000'
black = '#000000'

# Define file names
csvFileName = 'PCMMYM79013.csv'
resultsFileName = 'Results'


# PULL DATA FROM CSV FILE
def csvDataPull():
    rows = []
    with open(csvFileName, 'rb') as f:
        reader = csv.reader(f)
        for row in reader:
            for i, element in enumerate(row):
                try:
                    row[i] = float(row[i])
                except ValueError:
                    pass
            rows.append(row)
        return rows


# WRITES PULLED DATA TO A NEW EXCEL FILE
# This makes it easier to index the data base on the rows and columns
def datawriter(firstRow, dataArray):
    for i, element in enumerate(dataArray):
        column = 'A' + str(i + firstRow)
        worksheet.write_row(column, dataArray[i])


# GETS THE COLUMNS WITH THE INFORMATION FOR CALCULATIONS
def dataColumnLetters(columns):
    letters = []
    for i in range(0, columns):
        sLetter = xl_rowcol_to_cell(0, i)
        sLetter = sLetter[:-1]
        letters.append(sLetter)
    print(letters)
    return letters


def dataFiller(iRows):
    for i in sColumnLetters[9:]:
        row = []
        maxValCheck = ['Max Value']
        avgValCheck = ['Avg Value']
        minValCheck = ['Min Value']
        titles = ['Heading', 'Standard Deviation', 'Max Limit', 'Max Value', 'Average Value', 'Min value',
                  'Min Limit', 'Max Limit 75%', 'Mid Limit 50%', 'Min Limit 25%', '100%', '90%', '75%', 'Max Value %',
                  'Avg Value %', 'Min Value %', '25%', '10%', '0%']
        for j in range(5, 12):
            row.append(i + str(iRows + j))
            worksheet.write('A' + str(iRows + j), titles[j - 5])
        for j in range(14, 17):
            row.append(i + str(iRows + j))
            worksheet.write('A' + str(iRows + j), titles[j - 8])
        for j in range(19, 28):
            row.append(i + str(iRows + j))
            worksheet.write('A' + str(iRows + j), titles[j - 9])

        # Add excel formula's to output results
        headingFormula = '=' + i + str(1)
        stdDevFormula = '=STDEV(' + i + str(5) + ':' + i + str(19) + ')'
        maxLimit = '=' + i + str(4)
        minLimit = '=' + i + str(3)
        maxFormula = '=MAX(' + i + str(5) + ':' + i + str(19) + ')'
        avgFormula = '=AVERAGE(' + i + str(5) + ':' + i + str(19) + ')'
        minFormula = '=MIN(' + i + str(5) + ':' + i + str(19) + ')'
        maxLimPCT = '=((' + row[2] + '-' + row[6] + ') * 0.75) +' + row[6]
        minLimPCT = '=((' + row[2] + '-' + row[6] + ') * 0.25) +' + row[6]
        avgLimPCT = '=((' + row[2] + '-' + row[6] + ') * 0.5) +' + row[6]
        minValPCT = '=(((' + row[5] + '-' + row[6] + ') / (' + row[2] + '-' + row[6] + ')) * 100 )'
        avgValPCT = '=(((' + row[4] + '-' + row[6] + ') / (' + row[2] + '-' + row[6] + ')) * 100 )'
        maxValPCT = '=(((' + row[3] + '-' + row[6] + ') / (' + row[2] + '-' + row[6] + ')) * 100 )'
        pct75 = 75
        pct25 = 25
        pct90 = 90
        pct10 = 10
        pct100 = 100
        pct0 = 0
        formulas = [headingFormula, stdDevFormula, maxLimit, maxFormula,
                    avgFormula, minFormula, minLimit, maxLimPCT, avgLimPCT, minLimPCT, pct100, pct90, pct75, maxValPCT,
                    avgValPCT, minValPCT, pct25, pct10, pct0]
        for j in range(0, len(formulas)):
            worksheet.write(row[j], formulas[j])


# GRAPH THE DATA THAT HAS BEEN CALCULATED
def setupGraph(rowEnd):
    chart = workbook.add_chart({'type': 'line'})
    for i in range((rowEnd + 19), (rowEnd + 28)):
        lColors = ['red', 'orange', 'yellow', 'blue', 'green', 'purple', 'yellow', 'orange', 'red']
        color = lColors[i - (rowEnd + 19)]
        name = '==Sheet1!$A$' + str(i)
        values = '==Sheet1!$J$' + str(i) + ':$DA$' + str(i)
        chart.add_series({
            'name': name,
            'categories': '=Sheet1!$J$1:$DA$1',
            'values': values,
            'line': {'color': color, 'width': 1.5},
        })
    chart.set_x_axis({'major_gridlines': {
        'visible': True,
        'line': {'width': 1.25},
    },
        'interval_unit': 1,
        'label_position': 'low',
    })
    chart.set_y_axis({
        'min': -40,
        'max': 140,
    })
    return chart


def cellFormat(wb):
    format1 = wb.add_format({'bg_color': red,
                             'font_color': black})
    format2 = wb.add_format({'bg_color': yellow,
                             'font_color': black})
    format3 = wb.add_format({'bg_color': orange,
                             'font_color': black})
    formattedCols = sColumnLetters[9] + str(rows + 22) + ':' + sColumnLetters[-1] + str(rows + 24)
    worksheet.conditional_format(formattedCols, {
        'type': 'cell',
        'criteria': 'not between',
        'maximum': 100,
        'minimum': 0,
        'format': format1
    })
    worksheet.conditional_format(formattedCols, {
        'type': 'cell',
        'criteria': 'not between',
        'maximum': 90,
        'minimum': 10,
        'format': format3
    })
    worksheet.conditional_format(formattedCols, {
        'type': 'cell',
        'criteria': 'not between',
        'maximum': 75,
        'minimum': 25,
        'format': format2
    })


def outlierArray(array):
    newArray = []
    for i, element in enumerate(array):
        newArray.append(array[i][1])
    newArray = list(set(newArray))
    return newArray


# CREATE NEW EXCEL FILE FOR RESULTS
workbook = xlsxwriter.Workbook(resultsFileName + '.xlsx')
worksheet = workbook.add_worksheet()

dataRows = csvDataPull()
rows = len(dataRows)
cols = len(dataRows[0])
datawriter(1, dataRows)
sColumnLetters = dataColumnLetters(cols)

print (rows, cols)
under25 = []
under10 = []
over75 = []
over90 = []

for i in range(9, cols):
    colArray = []
    for j in range(4, rows):
        try:
            limitRange = float(dataRows[3][i]) - float(dataRows[2][i])
            limitRange25PCT = (limitRange * 0.25) + float(dataRows[2][i])
            limitRange75PCT = (limitRange * 0.75) + float(dataRows[2][i])
            limitRange90PCT = (limitRange * 0.90) + float(dataRows[2][i])
            limitRange10PCT = (limitRange * 0.10) + float(dataRows[2][i])

            curRow = j + 1
            if dataRows[j][i] > limitRange90PCT:
                over90.append([dataRows[j][4], dataRows[0][i]])
            elif dataRows[j][i] > limitRange75PCT:
                over75.append([dataRows[j][4], dataRows[0][i]])
            elif dataRows[j][i] < limitRange10PCT:
                under10.append([dataRows[j][4], dataRows[0][i]])
            elif dataRows[j][i] < limitRange25PCT:
                under25.append([dataRows[j][4], dataRows[0][i]])
        except:
            pass

over90Test = outlierArray(over90)
over75Test = outlierArray(over75)
under10Test = outlierArray(under10)
under25Test = outlierArray(under25)
print(over90)
print(over90Test)
print(over75)
print(over75Test)
print(under10)
print(under10Test)
print(under25)
print(under25Test)

worksheet2 = workbook.add_worksheet()
worksheet2.write('A1', 'Hello')

# GET ALL THE LETTER INDEX'S OF THE COLUMNS NEEDED TO BE ANALYSED, TO MAKE IT EASIER TO USE EXCEL FORMULAS
dataFiller(rows)
limitsChart = setupGraph(rows)
cellFormat(workbook)

chartPos = 'A' + str(rows + 30)
worksheet.insert_chart(chartPos, limitsChart, {'x_scale': 3, 'y_scale': 1.2})

workbook.close()
