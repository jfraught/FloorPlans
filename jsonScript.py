import json
import xlsxwriter
import os, os.path
from xlsxwriter.utility import xl_rowcol_to_cell

#Workbook

def makeWorkbook():
    #Set up Workbook
    workbook = xlsxwriter.Workbook('FloorPlan.xlsx')
    #TODO Set up Summary
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
        name = "FP" + str(number)
        number += 1
        makeWorksheet(name, workbook, floorPlanList, orthoWallsList, correctWallsList)
    #Wall Groups For Summary
         
    #Make Summary
   
    #Close workbook
    workbook.close()

def makeWorksheet(name, workbook, floorPlanList, orthoWallsList, correctWallsList):
    print(name)
    worksheet = workbook.add_worksheet(name)
    worksheet.set_column(0,0,29.5)
    percent_fmt = workbook.add_format({'num_format': '0.00%'})
    bold_percent_fmt = workbook.add_format({'num_format': '0.00%', 'bold': True})
    bold = workbook.add_format({'bold': True})
    formatExcel(worksheet, floorPlanList, orthoWallsList, correctWallsList, percent_fmt, bold, bold_percent_fmt)
   
def formatExcel(worksheet, floorPlanList, orthoWallsList, correctWallsList, percent_fmt, bold, bold_percent_fmt):
    #Room Count
    rooms = roomsCount(worksheet, floorPlanList)
    #Wall Count
    worksheet.write('A3', 'Wall Number')
    walls = wallCount(worksheet, orthoWallsList)
    #Walls
    worksheet.write('A4', 'Ortho Walls in Feet')
    displayWalls(orthoWallsList, 3, worksheet)
    worksheet.write('A5', 'Corrected Walls in Feet')
    displayWalls(correctWallsList, 4, worksheet)
    #Absolute Value Difference
    absoluteValueDifference(worksheet, orthoWallsList)
    #Percentage Difference, Contribution
    percentageDifference(worksheet, orthoWallsList, percent_fmt)
    weightedPercentage(worksheet, orthoWallsList, percent_fmt)
    contributionToWeight(worksheet, orthoWallsList)
    #Average Difference in Inches Wall Groups
    worksheet.write('A11', 'Average difference in inches - all walls', bold)
    worksheet.write_formula('B11', '=SUM(B6:CA6)/COUNT(B6:CA6)', bold)
    worksheet.write('A12', 'Average difference < 5 feet walls (in)')
    worksheet.write_formula('B12', '=SUMIFS(B6:CA6,B5:CA5,"<5")/COUNTIF(B5:CA5,"<5")')
    worksheet.write('A13', 'Average difference 5-15 feet walls (in)')
    worksheet.write_formula('B13', '=SUMIFS(B6:CA6,B5:CA5,">5",B5:CA5,"<15")/COUNTIFS(B5:CA5,">5",B5:CA5,"<15")')
    worksheet.write('A14', 'Average difference 15-25 feet walls (in)')
    worksheet.write_formula('B14', '=SUMIFS(B6:CA6,B5:CA5,">15",B5:CA5,"<25")/COUNTIFS(B5:CA5,">15",B5:CA5,"<25")')
    worksheet.write('A15', 'Average difference 25 > feet walls (in)')
    worksheet.write_formula('B15', '=SUMIFS(B6:CA6,B5:CA5,">25")/COUNTIF(B5:CA5,">25")')
    #Average Difference In % Wall Groups
    worksheet.write('A17','Average % difference - all walls', bold)
    worksheet.write_formula('B17', '=SUM(B7:CA7)/COUNT(B7:CA7)', bold_percent_fmt)
    worksheet.write('A18','Average % difference <5 feet walls')
    worksheet.write_formula('B18', '=SUMIFS(B7:CA7,B5:CA5,"<5")/COUNTIFS(B5:CA5,"<5")', percent_fmt)
    worksheet.write('A19','Average % difference 5-15 feet walls')
    worksheet.write_formula('B19', '=SUMIFS(B7:CA7,B5:CA5,">5",B5:CA5,"<15")/COUNTIFS(B5:CA5,">5",B5:CA5,"<15")', percent_fmt)
    worksheet.write('A20','Average % difference 15-25 feet walls')
    worksheet.write_formula('B20', '=SUMIFS(B7:CA7,B5:CA5,">15",B5:CA5,"<25")/COUNTIFS(B5:CA5,">15",B5:CA5,"<25")', percent_fmt)
    worksheet.write('A21','Average % difference >25 feet walls')
    worksheet.write_formula('B21', '=SUMIFS(B7:CA7,B5:CA5,">25")/COUNTIF(B5:CA5,">25")', percent_fmt)
    worksheet.write('A23','Weighted % Difference Average', bold)
    worksheet.write_formula('B23', '=SUM(B8:CA8)', bold_percent_fmt)

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
    #Average Weighted Percentage Difference
    worksheet.write('A5', 'Weighted Percentage Difference per Wall')
    percent_fmt = workbook.add_format({'num_format': '0.00%'})
    # Wall < 5
    worksheet.write('A6', 'Walls < 5ft Avg Difference per Wall')
    # Wall < 5 > 15
    worksheet.write('A7', '5ft < Walls < 15ft Avg Difference per Wall')
    # Wall < 15 > 25
    worksheet.write('A8', '15ft < Walls < 25ft Avg Difference per Wall')
    # Wall > 25
    worksheet.write('A9', '25ft < Walls Avg Difference per Wall')
    return 0



#Wall Functions

#def avgDifferenceInWallGroups():


        
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

#Good
def wallCount(worksheet, orthoWallsList):
    worksheet.write('A2', 'Wall Count')
    worksheet.write('B2', len(orthoWallsList))
    return len(orthoWallsList)

#Good 
def displayWalls(walls, row, worksheet):  
    col = 1
    index = 0
    for wall in walls:
        worksheet.write(2, col, index)
        worksheet.write(row, col, wall) 
        col += 1
        index += 1

#Good
def absoluteValueDifference(worksheet, orthoWallsList):
    col = 1
    worksheet.write(5, 0, 'Absolute Value Difference in Inches')
    for x in range(len(orthoWallsList)):
        cellFour = xl_rowcol_to_cell(4, col)
        cellThree = xl_rowcol_to_cell(3, col)
        formulaString = "=ABS(" + str(cellThree) + "-" + str(cellFour) + ") * 12"
        worksheet.write(5, col, formulaString)
        col += 1
#Good        
def percentageDifference(worksheet, orthoWallsList, percent_fmt):
    worksheet.write(6, 0, 'Percentage Difference')
    worksheet.set_row(6, None, percent_fmt)
    col = 1
    for x in range(len(orthoWallsList)):       
        cellFive = xl_rowcol_to_cell(5, col)
        cellThree = xl_rowcol_to_cell(3, col)
        formulaString = "=" + str(cellFive) + "/" + str(cellThree) + "/ 12" 
        worksheet.write(6, col, formulaString)
        col += 1
    
def contributionToWeight(worksheet, orthoWallsList):
    col = 1
    worksheet.write('A9', 'Contribution to weight')
    for x in range(len(orthoWallsList)):
        cell = xl_rowcol_to_cell(3, col)
        formulaString = "=" + str(cell) + "/SUMIF(B6:CA6,\"<>0\",B4:CA4)"
        worksheet.write_formula(8, col, formulaString)
        col += 1


def weightedPercentage(worksheet, orthoWallsList, percent_fmt):
    worksheet.write('A8', 'Weighted Percentage')
    col = 1
    worksheet.set_row(7, None, percent_fmt)
    for x in range(len(orthoWallsList)):
        cellEight = xl_rowcol_to_cell(8, col) 
        cellSix =  xl_rowcol_to_cell(6, col)
        formulaString = "=" + str(cellEight) + "*" + str(cellSix)
        worksheet.write_formula(7, col, formulaString)
        col += 1

def averageDifference(worksheet):
    worksheet.write('A9', 'Average difference in Inches')
    worksheet.write_formula('B9', '=SUM(B5:CA5)/COUNTIF(B5:CA5, ">0")')

def metersTofFeet(meters):
    feet = meters / 0.3048
    return feet

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
