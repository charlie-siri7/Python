# imports for sys, pandas, math
import sys
import pandas as pd
import math
import os
import matplotlib.pyplot as plt

def writeLine(f, text):
    f.write(text)
    f.write("\n")

def printTableHeader(f, header):
    writeLine(f, "<!DOCTYPE html>")
    writeLine(f, "<html>")
    writeLine(f, "<head>")
    writeLine(f, "<style>")
    writeLine(f, "td {")
    writeLine(f, "border: 1px solid black;")
    writeLine(f, "text-align: center;")
    writeLine(f, "}")
    writeLine(f, "</style>")
    writeLine(f, "</head>")
    writeLine(f, "<body>")
    writeLine(f, "<h1>" + header + "</h1>")
    writeLine(f, "<table>")

if os.path.exists("CardStats.html"):
    os.remove("CardStats.html")

# read excel file and print its contents
data = pd.read_excel("CardStats.xlsx")
print(data)

columnNames = list(data.columns)
allNamesList = data['CARD'].values.tolist()

f = open("CardStats.html", "a")
printTableHeader(f, "Card Stats")

writeLine(f, "<tr>")
for name in columnNames:
    writeLine(f, "<th>" + name + "</th>")
writeLine(f, "</tr>")

for i in range(len(allNamesList)):
    writeLine(f, "<tr>")
    for j in range(len(columnNames)):
        f.write("<td>")
        f.write(str(data[columnNames[j]].values.tolist()[i]))
        f.write("</td>\n")
    writeLine(f, "</tr>")

writeLine(f, "</table>")
writeLine(f, "</body>")
writeLine(f, "</html>")
f.close()

running = True
while running == True:
    # ex 1: sort_increasing: dps include: damage seconds
    # ex 2: get: dps damage hit_speed of knight
    # ex 3: scatterplot damage hit_speed
    command = input("Enter command: (ex: 'sort_increasing: dps include: damage seconds'): ")
    tokens = command.split(" ")
    mode = "sort"
    if (tokens[0].upper() == 'SORT_INCREASING:'):
        sortMode = 0
    elif (tokens[0].upper() == 'SORT_DECREASING:'):
        sortMode = 1
    elif (tokens[0].upper() == 'GET:'):
        mode = "card"
    else:
        mode = "scatterplot"
    
    if (mode == "sort"):
        column = tokens[1].upper()
        extraColumnNames = []
        i = 3
        while i < len(tokens):
            extraColumnNames.append(tokens[i])
            i += 1

        # stores the list of the values of each extra column (default sorting) in a list of lists
        oldExtraColumns = []
        for i in extraColumnNames:
            oldExtraColumns.append(data[i.upper()].values.tolist())

        # declare list of column to sort, declare variables such as other lists, counters
        columnList = data[column].values.tolist()
        sortedList = []
        indexList = []
        nameList = []
        numeric = True
        stringCounter = 0
        nanCounter = 0

        # set the counts of the number of strings and non-numbers in the column to sort
        for i in range(len(columnList)):
            if type(columnList[i]) is str:
                stringCounter += 1
            elif math.isnan(columnList[i]):
                print(type(columnList[i]))
                nanCounter += 1

        # set numeric to false is the total of strings, non-numbers equals the length of the column to sort
        if len(columnList) == (stringCounter + nanCounter):
            numeric = False

        if numeric:
            if sortMode == 0:
                # if items are numeric and to be sorted from smallest to largest:

                # set all unfilled values of column list to the max int value
                for i in range(len(columnList)):
                    if math.isnan(columnList[i]):
                        columnList[i] = sys.maxsize
                
                # runs while column list still has items and the min value isn't the highest int
                while len(columnList) > 0 and min(columnList) < sys.maxsize:
                    # add the smallest columnList item to sortedList, append its index to indexList, set the recently added item in columnList to the max int value
                    sortedList.append(min(columnList))
                    indexList.append(columnList.index(min(columnList)))
                    columnList[columnList.index(min(columnList))] = sys.maxsize
            else:
                # if items are numeric and to be sorted from largest to smallest:

                # set all unfilled values of column list to -1
                for i in range(len(columnList)):
                    if math.isnan(columnList[i]):
                        columnList[i] = -1

                # runs while column list still has items and the max value isn't -1
                while len(columnList) > 0 and max(columnList) > -1:
                    # add the largest columnList item to sortedList, append its index to indexList, set the recently added item in columnList to -1
                    sortedList.append(max(columnList))
                    indexList.append(columnList.index(max(columnList)))
                    columnList[columnList.index(max(columnList))] = -1
        else:
            # initialize counter for number of items found
            found = 0
            if sortMode == 0:
                # if to be sorted from A-Z alphabetically, run until all strings are found
                while found < stringCounter:
                    min = ""
                    minIndex = 0
                    # loop through the column list to find the lowest value alphabetically
                    for i in range(len(columnList)):
                        if ((type(columnList[i]) is str) == True) and columnList[i] != "":
                            if min == "" or columnList[i] <= min:
                                min = columnList[i]
                                minIndex = i
                    # append lowest value to sorted list and its index to index list, clear value in column list, increment found items counter
                    sortedList.append(min)
                    indexList.append(minIndex)
                    columnList[minIndex] = ""
                    found += 1
            else:
                # if to be sorted from Z-A alphabetically, run until all strings are found
                while found < stringCounter:
                    max = ""
                    maxIndex = 0
                    # loop through the column list to find the highest value alphabetically
                    for i in range(len(columnList)):
                        if ((type(columnList[i]) is str) == True) and columnList[i] != "":
                            if max == "" or columnList[i] >= max:
                                max = columnList[i]
                                maxIndex = i
                    # append highest value to sorted list and its index to index list, clear value in column list, increment found items counter
                    sortedList.append(max)
                    indexList.append(maxIndex)
                    columnList[maxIndex] = ""
                    found += 1

        # Add the name of each card in newly sorted order to a list of names
        for i in indexList:
            nameList.append(data['CARD'].values.tolist()[i])

        # new list whose elements will be lists representing each extra column to be displayed
        newExtraColumns = []
        for name in extraColumnNames:
            # For each column name, append the item under the column at each index of indexList to a new list called newExtraColumn
            newExtraColumn = []
            for index in indexList:
                item = str(data[name.upper()].values.tolist()[index])
                if item.upper() == "NAN":
                    item = "N/A"
                newExtraColumn.append(item)
            # newExtraColumn list should be a sorted additional column whose values map to that of the newly sorted column, so append it to newExtraColumns
            newExtraColumns.append(newExtraColumn)

        # make a dataframe with the name of cards and sorted column
        subsetDataFrame = pd.DataFrame({'NAME': nameList, column: sortedList})
        # append each extra column to the dataframe
        for i in range(len(extraColumnNames)):
            subsetDataFrame.insert(len(subsetDataFrame.columns), extraColumnNames[i].upper(), newExtraColumns[i])
        # print sorted dataframe with context
        print("Sorted by", column)
        print(subsetDataFrame)

        columnNames = []
        columnNames.append("CARD")
        columnNames.append(column.upper())
        for name in extraColumnNames:
            columnNames.append(name.upper())

        pageName = column + ".html"
        pageTitle = "Cards sorted by " + column

        if os.path.exists(pageName):
            os.remove(pageName)

        f = open(pageName, "a")
        printTableHeader(f, pageTitle)

        writeLine(f, "<tr>")
        writeLine(f, "<th>CARD</th>")
        writeLine(f, "<th>" + column + "</th>")
        for name in extraColumnNames:
            writeLine(f, "<th>" + name.upper() + "</th>")
        writeLine(f, "</tr>")

        for i in indexList:
            writeLine(f, "<tr>")
            for j in range(len(columnNames)):
                f.write("<td>")
                item = str(data[columnNames[j]].values.tolist()[i])
                if item.upper() == "NAN":
                    item = "N/A"
                f.write(item)
                f.write("</td>\n")
            writeLine(f, "</tr>")

        writeLine(f, "</table>")
        writeLine(f, "</body>")
        writeLine(f, "</html>")
        f.close()
    elif mode == "card":
        attributes = ["NAME"]
        values = []
        i = 1
        while tokens[i].upper() != "OF":
            attributes.append(tokens[i].upper())
            i += 1
        card = tokens[i + 1].lower()
        values.append(card)
        index = data['CARD'].values.tolist().index(card)
        for i in range(len(attributes) - 1):
            values.append(data[attributes[i + 1].upper()].values.tolist()[index])
        cardDataFrame = pd.DataFrame({'Attribute: ': attributes, 'Value: ': values})
        print(cardDataFrame)

        pageName = card.upper() + "_STATS.html"
        pageTitle = "Specified " + card + " stats"
        if os.path.exists(pageName):
            os.remove(pageName)
        f = open(pageName, "a")
        printTableHeader(f, pageTitle)
        writeLine(f, "<tr>")
        writeLine(f, "<th>Attribute:</th>")
        writeLine(f, "<th>Value:</th>")
        writeLine(f, "</tr>")
        for i in range(len(attributes)):
            writeLine(f, "<tr>")
            writeLine(f, "<th>" + str(attributes[i]) + "</th>")
            writeLine(f, "<th>" + str(values[i]) + "</th>")
            writeLine(f, "</tr>")
        writeLine(f, "</table>")
        writeLine(f, "</body>")
        writeLine(f, "</html>")
        f.close()
    elif mode == "scatterplot":
        xlist = data[tokens[1].upper()].tolist()
        ylist = data[tokens[2].upper()].tolist()
        plt.scatter(xlist, ylist)
        plt.show()

    runningStr = input("Enter 'Y' to run the program again: ")
    if (runningStr != 'Y'):
        running = False