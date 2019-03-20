import json
import xlsxwriter
import os, os.path





#Workbook

#Get JSON Function
#Loop through folder of JSON files and add a worksheet to the workbook for each. 

def getFilePath():
    directory = '/Users/i25203/Desktop/JSON'
    files = []
    dirLength = 0
    fileName = ''
    for dirpath, dirname, filenames in os.walk(directory):
        dirLength = (len(filenames))
    for x in range(dirLength):
        fileName = directory + "/" + filenames[x]
        file = getJSONFile(fileName)
        files.append(file)
    return files

def getJSONFile(path):
    jsonFloorPlan = ''
    with open(path, 'r') as f:
        f_contents = f.read()
        jsonFloorPlan = f_contents
    return jsonFloorPlan



#Make Excel Function

def makeWorkbook():
    #Set up Workbook
    workbook = xlsxwriter.Workbook('FloorPlan.xlsx')
    #loop through files
    files = getFilePath()
    number = 1
    floorPlans = 1
    for file in files:
        #Worksheet set up
        floorPlanList = makeFloorPlanList(file)
        floorPlans += 1
        orthoWallsList = getWalls("orthorectified", floorPlanList)
        correctWallsList = getWalls("correctedMeasurment", floorPlanList)  
        if len(orthoWallsList) != len(correctWallsList):
            correctWallsList = fixCorrectedWallsList(correctWallsList)

        groupWalls(orthoWallsList, correctWallsList)
        differenceSum = wallsByHand(orthoWallsList, correctWallsList)
        if differenceSum != 0:
            name = "FP" + str(number)
            number += 1
            makeWorksheet(name, workbook)
    #Close workbook
    workbook.close()

def makeWorksheet(name, workbook):
    #Remember TODO handle floor plans with no differences. 
    worksheet = workbook.add_worksheet(name)
    worksheet.set_column(0,0,27)
    percent_fmt = workbook.add_format({'num_format': '0.00%'})
    formatExcel(worksheet)
    #Locacl Variables


def formatExcel(worksheet):
    roomsCount(worksheet)
    wallCount(worksheet)
    worksheet.write('A3', 'Ortho Walls in Feet')
    displayWalls(0, tupleWallList, 2, worksheet)
    worksheet.write('A4', 'Corrected Walls in Feet')
    displayWalls(1, tupleWallList, 3, worksheet)
    absoluteValueDifference(worksheet)
    percentageDifference(worksheet)
    averageDifference(worksheet)
    contributionToWeight(worksheet)
    weightedPercentage(worksheet)
    

def makeFloorPlanList(jsonPlan):
    pyFloorPlan = json.loads(jsonPlan)
    floorPlanList = pyFloorPlan["floorPlans"]
    return floorPlanList

#Wall Functions

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

def wallCount(worksheet):
    worksheet.write('A2', 'Wall Count')
    worksheet.write('B2', len(orthoWallsList))

def groupWalls(ortho, correct):
    index = 0
    for wall in ortho:
        walls = (wall, correct[index])
        index += 1
        tupleWallList.append(walls)
    tupleWallList.sort(key = sortOrtho)

def sortOrtho(val):
    return val[0]

def displayWalls(wallType, walls, row, worksheet):  
    col = 1
    for wall in walls:
        worksheet.write(row, col, wall[wallType])
        col += 1
        
def absoluteValueDifference(worksheet):
    col = 1
    worksheet.write(4, 0, 'Absolute Value Difference in Inches')
    for x in range(len(tupleWallList)):
        tupleAtIndex = tupleWallList[x]
        orthoWall = tupleAtIndex[0]
        correctWall = tupleAtIndex[1]
        differenceInFeet = abs(orthoWall - correctWall)
        differenceInInches = feetToInches(differenceInFeet)
        differenceList.append(differenceInFeet)
        worksheet.write(4, col, differenceInInches)
        col += 1
        
def percentageDifference(worksheet):
    worksheet.write(5, 0, 'Percentage Difference')
    worksheet.set_row(5, None, percent_fmt)
    col = 1
    for x in range(len(tupleWallList)):
        tupleAtIndex = tupleWallList[x]
        orthoWall = tupleAtIndex[0]
        percent = differenceList[x] / orthoWall
        percentageList.append(percent)
        worksheet.write(5, col, percent)
        col += 1

def averageDifference(worksheet):
    wallsByHand = 0
    differenceSum = 0
    for value in differenceList:
        if value != 0:
            wallsByHand += 1
            differenceSum += value
    difference = differenceSum / wallsByHand
    differenceInInches = feetToInches(difference)
    worksheet.write('A7', 'Average difference in Inches')
    worksheet.write(6, 1, differenceInInches)

def contributionToWeight(worksheet):
    orthoSum = 0
    contribution = 0
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
            worksheet.write(7, col, contribution)
        else: 
            contribution = 0
        contributionList.append(contribution)
        col += 1

def weightedPercentage(worksheet):
    worksheet.write('A9', 'Weighted Percentage')
    worksheet.write('A10', 'Weighted Difference Average')
    weightedPercentage = 0
    weightedPercentageAverage = 0
    col = 1
    listCount = 0
    worksheet.set_row(8, None, percent_fmt)
    worksheet.set_row(9, None, percent_fmt)
    for x in range(len(percentageList)):
        if percentageList[x] != 0:
            weightedPercentage = contributionList[x] * percentageList[x] 
            listCount += 1
        else:
            weightedPercentage = 0
        worksheet.write(8, col, weightedPercentage)
        col += 1
        weightedPercentageAverage += weightedPercentage
    worksheet.write('B10', weightedPercentageAverage)

def metersTofFeet(meters):
    feet = meters / 0.3048
    return feet

def feetToInches(feet):
    inches = feet * 12
    return inches

#Room Functions

def roomsCount(worksheet):
    roomList = []
    worksheet.write('A1', 'Room Count')
    roomNumber = 1
    print(floorPlanList)
    for floor in floorPlanList:
        if floor["type"] == "orthorectified":
            rooms = floor["rooms"]
            for room in rooms:
                roomNumber += 1
                roomList.append(room)
    worksheet.write(0, 1, len(roomList))

#Global
tupleWallList = []
differenceList = []
percentageList = [] 
contributionList = []
weightedPercentagList = []
floorPlanList = []
orthoWallsList = []
correctWallsList = [] 
percent_fmt = ''

makeWorkbook()

