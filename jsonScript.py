import json
import xlsxwriter
import os, os.path
from xlsxwriter.utility import xl_rowcol_to_cell

#Workbook

def makeWorkbook():
    #Set up Workbook
    workbook = xlsxwriter.Workbook('FloorPlan.xlsx')
    #Set up Summary Worksheet
    averageDifferenceList = []
    weightedPercentageList = []
    totalRoomCount = 0
    totalWallCount = 0 
    wallGroupsDifferenceList = []
    wallsAvg25List = []
    wallsAvg15List = []
    wallsAvg5List = []
    wallsAvg0List = []
    #loop through files
    files = getFilePath()
    number = 1
    floorPlans = 0
    for file in files:
        #Worksheet set up
        floorPlanList = makeFloorPlanList(file)
        floorPlans += 1
        orthoWallsList = getWalls("orthorectified", floorPlanList)
        correctWallsList = getWalls("correctedMeasurment", floorPlanList)  
        if len(orthoWallsList) != len(correctWallsList):
            correctWallsList = fixCorrectedWallsList(correctWallsList)
        if len(correctWallsList) != 0:
            tupleWallList = groupWalls(orthoWallsList, correctWallsList)
            differenceSum = wallsByHand(orthoWallsList, correctWallsList)
            if differenceSum != 0:
                name = "FP" + str(number)
                number += 1
                summaryList = makeWorksheet(name, workbook, floorPlanList, orthoWallsList, tupleWallList)
                averageDifferenceList.append(summaryList[0])
                weightedPercentageList.append(summaryList[1])
                totalRoomCount += summaryList[2]
                totalWallCount += summaryList[3]
                #Wall Groups For Summary
                wallGroupsDifferenceList = differenceInWallGroups(tupleWallList)
                wallsGroupsAvgDifferenceList = avgDifferenceInWallGroups(wallGroupsDifferenceList)
                wallsAvg25List.append(wallsGroupsAvgDifferenceList[0])
                wallsAvg15List.append(wallsGroupsAvgDifferenceList[1])
                wallsAvg5List.append(wallsGroupsAvgDifferenceList[2])
                wallsAvg0List.append(wallsGroupsAvgDifferenceList[3])
    groupsAvgListsList = [wallsAvg0List, wallsAvg5List, wallsAvg15List, wallsAvg25List]
    #Make Summary
    summaryWorksheet(workbook, totalRoomCount, totalWallCount, weightedPercentageList, averageDifferenceList, floorPlans, groupsAvgListsList)
    #Close workbook
    workbook.close()

def makeWorksheet(name, workbook, floorPlanList, orthoWallsList, tupleWallList):
    print(name)
    worksheet = workbook.add_worksheet(name)
    worksheet.set_column(0,0,27)
    percent_fmt = workbook.add_format({'num_format': '0.00%'})
    summaryList = formatExcel(worksheet, floorPlanList, orthoWallsList, tupleWallList, percent_fmt)
    return summaryList

def formatExcel(worksheet, floorPlanList, orthoWallsList, tupleWallList, percent_fmt):
    rooms = roomsCount(worksheet, floorPlanList)
    walls = wallCount(worksheet, orthoWallsList)
    worksheet.write('A3', 'Ortho Walls in Feet')
    displayWalls(0, tupleWallList, 2, worksheet)
    worksheet.write('A4', 'Corrected Walls in Feet')
    displayWalls(1, tupleWallList, 3, worksheet)
    differenceList = absoluteValueDifference(worksheet, tupleWallList)
    percentageList = percentageDifference(worksheet, tupleWallList, differenceList, percent_fmt)
    contributionList = contributionToWeight(worksheet, tupleWallList, differenceList)
    avgDif = averageDifference(worksheet, tupleWallList, differenceList)
    wghtAvrPer = weightedPercentage(worksheet, percentageList, contributionList, percent_fmt)
    summaryList = [avgDif, wghtAvrPer, rooms, walls]
    return summaryList

def summaryWorksheet(workbook, totalRoomCount, totalWallCount, weightedPercentageList, averageDifferenceList, floorPlans, groupsAvgListsList):
    worksheet = workbook.add_worksheet('Summary')
    worksheet.set_column(0,0,32)
    #Number of Floor Plans
    worksheet.write('A1', 'Number of Floor Plans')
    worksheet.write('B1', floorPlans)
    #Number of Walls
    worksheet.write('A2', 'Number of Walls')
    worksheet.write('B2', totalWallCount)
    #Number of Rooms
    worksheet.write('A3', 'Number of Rooms')
    worksheet.write('B3', totalRoomCount)
    #Average Difference in Inches
    worksheet.write('A4', 'Average Differnce in Inches per Wall')
    averageDifference = average(averageDifferenceList)
    worksheet.write('B4', round(averageDifference, 2))
    #Average Weighted Percentage Difference
    worksheet.write('A5', 'Weighted Percentage Difference per Wall')
    weightedPercentage = average(weightedPercentageList)
    percent_fmt = workbook.add_format({'num_format': '0.00%'})
    worksheet.set_row(4, None, percent_fmt)
    worksheet.write('B5', weightedPercentage)
    # Wall < 5
    worksheet.write('A6', 'Walls < 5ft Avg Difference per Wall')
    avg0 = average(groupsAvgListsList[0])
    worksheet.write('B6', round(avg0, 2))
    # Wall < 5 > 15
    worksheet.write('A7', '5ft < Walls < 15ft Avg Difference per Wall')
    avg5 = average(groupsAvgListsList[1])
    worksheet.write('B7', round(avg5, 2))
    # Wall < 15 > 25
    worksheet.write('A8', '15ft < Walls < 25ft Avg Difference per Wall')
    avg15 = average(groupsAvgListsList[2])
    worksheet.write('B8', round(avg15, 2))
    # Wall > 25
    worksheet.write('A9', '25ft < Walls Avg Difference per Wall')
    avg25 = average(groupsAvgListsList[3])
    worksheet.write('B9', round(avg25, 2))
    return 0

def average(x):
    return sum(x) / len(x)

#Wall Functions
def avgDifferenceInWallGroups(wallGroupsDifferenceList):
    avgDifferenceList = []
    avgDifference = 0
    for differenceList in wallGroupsDifferenceList:
        if len(differenceList) != 0:
           avgDifference = average(differenceList)
        else:
            avgDifference = 0
        avgDifferenceList.append(avgDifference)
    return avgDifferenceList

def differenceInWallGroups(tupleWallList):
    greaterThan25List = []
    greaterThan15List = []
    greaterThan5List = []
    lessThan5List = []
    for x in range(len(tupleWallList)):
        tupleAtIndex = tupleWallList[x]
        orthoWall = tupleAtIndex[0]
        correctedWall = tupleAtIndex[1]
        orthoFeet = metersTofFeet(orthoWall)
        correctedFeet = metersTofFeet(correctedWall)
        difference = abs(orthoFeet - correctedFeet)
        differenceInInches = feetToInches(difference)
        if orthoFeet > 25:
            greaterThan25List.append(differenceInInches)
        elif orthoFeet > 15:
            greaterThan15List.append(differenceInInches)
        elif orthoFeet > 5:
            greaterThan5List.append(differenceInInches)
        else:   
            lessThan5List.append(differenceInInches)
    print(len(greaterThan25List))
    return [greaterThan25List, greaterThan15List, greaterThan5List, lessThan5List]

def wallsByHand(ortho, correct):
    totalDifference = 0
    for x in range(len(ortho)):
        difference = abs(ortho[x] - correct[x])
        totalDifference += difference
    return totalDifference
        
def getWalls(type, floorPlanList): 
    wallsList = []
    for floor in floorPlanList:
        if floor["type"] == type:
            walls = floor["walls"]
            for wall in walls:
                lengthInMeters = wall["length"]
                lengthInFeet = metersTofFeet(lengthInMeters)
                wallsList.append(lengthInFeet)
    return wallsList

def fixCorrectedWallsList(correctedWallsList):
    halfListCount = len(correctedWallsList) / 2
    listCount = len(correctedWallsList)
    while listCount > halfListCount:
        del correctedWallsList[listCount - 1]
        listCount -= 1
    return correctedWallsList

def wallCount(worksheet, orthoWallsList):
    worksheet.write('A2', 'Wall Count')
    worksheet.write('B2', len(orthoWallsList))
    return len(orthoWallsList)

def groupWalls(ortho, correct):
    index = 0
    tupleWallList = []
    for wall in ortho:
        walls = (wall, correct[index])
        index += 1
        tupleWallList.append(walls)
    tupleWallList.sort(key = sortOrtho)
    return tupleWallList

def sortOrtho(val):
    return val[0]

def displayWalls(wallType, walls, row, worksheet):  
    col = 1
    for wall in walls:
        worksheet.write(row, col, wall[wallType])
        col += 1
        
def absoluteValueDifference(worksheet, tupleWallList):
    col = 1
    differenceList = []
    worksheet.write(4, 0, 'Absolute Value Difference in Inches')
    for x in range(len(tupleWallList)):
        tupleAtIndex = tupleWallList[x]
        orthoWall = tupleAtIndex[0]
        correctWall = tupleAtIndex[1]
        differenceInFeet = abs(orthoWall - correctWall)
        differenceInInches = feetToInches(differenceInFeet)
        differenceList.append(differenceInFeet)
        cellFour = xl_rowcol_to_cell(3, col)
        cellThree = xl_rowcol_to_cell(2, col)
        formulaString = "=ABS(" + str(cellThree) + "-" + str(cellFour) + ") * 12"
        worksheet.write(4, col, formulaString)
        col += 1
    return differenceList
        
def percentageDifference(worksheet, tupleWallList, differenceList, percent_fmt):
    worksheet.write(5, 0, 'Percentage Difference')
    worksheet.set_row(5, None, percent_fmt)
    percentageList = []
    col = 1
    for x in range(len(tupleWallList)):
        tupleAtIndex = tupleWallList[x]
        orthoWall = tupleAtIndex[0]
        percent = differenceList[x] / orthoWall
        percentageList.append(percent)
        if percent != 0:
            cellFive = xl_rowcol_to_cell(4, col)
            cellThree = xl_rowcol_to_cell(2, col)
            formulaString = "=" + str(cellFive) + "/" + str(cellThree) + "/ 12" 
            worksheet.write(5, col, formulaString)
        col += 1
    return percentageList
    
def averageDifference(worksheet, tupleWallList, differenceList):
    wallsByHand = 0
    differenceSum = 0
    for value in differenceList:
        if value != 0:
            wallsByHand += 1
            differenceSum += value
    difference = differenceSum / wallsByHand
    differenceInInches = feetToInches(difference)
    worksheet.write('A9', 'Average difference in Inches')
    worksheet.write_formula('B9', '=SUM(B5:CA5)/COUNTIF(B5:CA5, ">0")')
    return differenceInInches

def contributionToWeight(worksheet, tupleWallList, differenceList):
    orthoSum = 0
    contribution = 0
    contributionList = []
    col = 1
    worksheet.write('A8', 'Contribution to weight')
    for x in range(len(tupleWallList)):
        tupleAtIndex = tupleWallList[x]
        orthoWall = tupleAtIndex[0]
        if differenceList[x] != 0:
            orthoSum += orthoWall
    for x in range(len(tupleWallList)):
        tupleAtIndex = tupleWallList[x]
        orthoWall = tupleAtIndex[0]
        difference = differenceList[x]
        if difference != 0.0:
            contribution = orthoWall / orthoSum
            cell = xl_rowcol_to_cell(2, col)
            formulaString = "=" + str(cell) + "/SUMIF(B5:CA5,\"<>0\",B3:CA3)"
            worksheet.write_formula(7, col, formulaString)
        else: 
            contribution = 0
        contributionList.append(contribution)
        col += 1
    return contributionList

def weightedPercentage(worksheet, percentageList, contributionList, percent_fmt):
    worksheet.write('A7', 'Weighted Percentage')
    worksheet.write('A10', 'Weighted Difference Average')
    weightedPercentage = 0
    weightedPercentageAverage = 0
    col = 1
    listCount = 0
    worksheet.set_row(6, None, percent_fmt)
    worksheet.set_row(9, None, percent_fmt)
    for x in range(len(percentageList)):
        if percentageList[x] != 0:
            weightedPercentage = contributionList[x] * percentageList[x] 
            listCount += 1
            cellEight = xl_rowcol_to_cell(7, col) 
            cellSix =  xl_rowcol_to_cell(5, col)
            formulaString = "=" + str(cellEight) + "*" + str(cellSix)
            worksheet.write_formula(6, col, formulaString)
        else:
            weightedPercentage = 0
        col += 1
        weightedPercentageAverage += weightedPercentage
    worksheet.write_formula('B10', '=SUM(B7:CA7)')
    return weightedPercentageAverage

def metersTofFeet(meters):
    feet = meters / 0.3048
    return feet

def feetToInches(feet):
    inches = feet * 12
    return inches

#Room Functions

def roomsCount(worksheet, floorPlanList):
    roomList = []
    worksheet.write('A1', 'Room Count')
    roomNumber = 1
    for floor in floorPlanList:
        if floor["type"] == "orthorectified":
            rooms = floor["rooms"]
            for room in rooms:
                roomNumber += 1
                roomList.append(room)
    worksheet.write(0, 1, len(roomList))
    return len(roomList)

#JSON Functions

def getFilePath():
    directory = '/Users/i25203/Desktop/JSON'
    files = []
    dirLength = 0
    fileName = ''
    for dirpath, dirname, filenames in os.walk(directory):
        dirLength = (len(filenames))
    for x in range(dirLength):
        fileName = directory + "/" + filenames[x]
        print(fileName)
        file = getJSONFile(fileName)
        files.append(file)
    return files

def getJSONFile(path):
    jsonFloorPlan = ''
    with open(path, 'r') as f:
        f_contents = f.read()
        jsonFloorPlan = f_contents
    return jsonFloorPlan

def makeFloorPlanList(jsonPlan):
    pyFloorPlan = json.loads(jsonPlan)
    floorPlanList = pyFloorPlan["floorPlans"]
    return floorPlanList

#To remove ds.store files after placing new group of logs in folder run "find . -name '.DS_Store' -type f -delete" in terminal. 

makeWorkbook()

