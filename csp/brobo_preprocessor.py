from buildXlsFile import buildXFile

def buildBroboProgram(data, bladeKerf, jobNumber, savePath, programNum = 1, fileName = None):
    broboSticksData = buildBroboCutData(data, bladeKerf)
    # print(f"BroboSticksData: {broboSticksData}")          #Used For Debugging
    tempProgramNum = programNum
    for stickData in broboSticksData:
        xlsFile = buildXFile(stickData, tempProgramNum, jobNumber, savePath, fileName)
        xlsFile.buildSheet()
        tempProgramNum += 1

#Build Brobo Cut Data
def buildBroboCutData(data, bladeKerf):
    data,bladeKerf = scaleData(data, bladeKerf)
    demandData = [[]]
    for stick in data:
        tempStick = []
        print (f"stick: {stick}")
        stick = orderList(stick)                                                   #TODO: There is a Bug here, and needs to fix
        tempStick.append(buildInitialAbsVal(stick, bladeKerf))
        stick = drop_first(stick)
        for measurment in stick:
            tempStick.append(ensure_negative(measurment))
        demandData.append(tempStick)
    return demandData

def scaleData(data, bladeKerf):
    bladeKerf = round(inchToDecimalInch(bladeKerf))
    for stick in data:
        for i, element in enumerate(stick):
            stick[i] = round(inchToDecimalInch(element))
    return data, bladeKerf

def buildInitialAbsVal(stickData, bladeKerf):
    kerfCount = total_kerf_count(stickData)
    kerfLength = total_kerf_length(kerfCount, bladeKerf)
    partsLength = total_parts_length(stickData)
    return abs(kerfLength + partsLength)

#Helper Functions
def inchToDecimalInch(measurement):
    return float(measurement) * 1000

def orderList(list):
    return sorted(list)

def total_parts_length(list):
    return sum(list)

def total_kerf_count(list):
    return (len(list) - 1)

def total_kerf_length(KerfCount, KerfWidth):
    return (KerfCount * KerfWidth)

def ensure_negative(measurement):
    return (abs(measurement) * -1)

def drop_first(list):
    return list[1:]