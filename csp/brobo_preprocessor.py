from buildXlsFile import buildXFile
from solver_handler import CuttingParameters
def buildBroboProgram(cutParams):
    solution = cutParams.getSolution()
    bladeKerf = cutParams.getBladeWidth()
    jobNumber = cutParams.getJobNumber()
    savePath = cutParams.getDirPath()
    programNum = 1
    fileName = None

    #######################################################################
    #TODO: We are rebuilding this function to use a passed in CuttingParameters object so that we can have more control over what is passed to the xls file builder.
    #######################################################################

    """
    Builds a Brobo program by processing the raw cut data and creating an Excel file for each stick.

    This function processes the raw cut data, prepares it for the Brobo machine, 
    it then creates an Excel file for each stick using the buildXlsFile Library. 
    The Excel files are saved in the specified save path with an optional file name.

    Args:
        data (list -> lists -> float): The raw cut data, where each inner list represents a stick and contains measurements.
        bladeKerf (float): The width of the saw blade.
        jobNumber (int): The job number.
        savePath (str): The path where the Excel files will be saved.
        programNum (int, optional): The starting program number. Defaults to 1.
        fileName (str, optional): The name of the Excel files. If None, a name will be auto generated.

    Dependencies:
        buildXlsFile.py

    Usage:
        buildBroboProgram(data, bladeKerf, jobNumber, savePath, programNum = 1, fileName = None)
    """
    broboSticksData = _buildBroboCutData(data, bladeKerf)
    tempProgramNum = programNum
    for stickData in broboSticksData:
        xlsFile = buildXFile(stickData, tempProgramNum, jobNumber, savePath, fileName)
        xlsFile.buildSheet()
        tempProgramNum += 1


############ Private Functions ############
        
def _buildBroboCutData(data, bladeKerf):
    """
    Builds the cut data for Brobo machine.

    Args:
        data (list): The input data.
        bladeKerf (float): The blade kerf value.

    Returns:
        list: The processed demand data.
    """
    data, bladeKerf = _scaleData(data, bladeKerf)
    demandData = []
    for stick in data:
        tempStick = []
        stick = _orderList(stick)
        tempStick.append(_buildInitialAbsVal(stick, bladeKerf))
        stick = _drop_first(stick)
        for measurement in stick:
            tempStick.append(_ensure_negative(measurement))
        demandData.append(tempStick)
    return demandData

def _scaleData(data, bladeKerf):
    """
    Scales the data and blade kerf to decimal inches.

    Parameters:
    data (list): The data to be scaled.
    bladeKerf (float): The blade kerf value.

    Returns:
    tuple: A tuple containing the scaled data and blade kerf.
    """
    bladeKerf = round(_inchToDecimalInch(bladeKerf))
    for stick in data:
        for i, element in enumerate(stick):
            stick[i] = round(_inchToDecimalInch(element))
    return data, bladeKerf

def _buildInitialAbsVal(stickData, bladeKerf):
    """
    Calculates the initial absolute value by adding the total kerf length and total parts length.

    Parameters:
    stickData (list): List of stick data.
    bladeKerf (float): Blade kerf value.

    Returns:
    float: The initial absolute value.
    """
    kerfCount = _total_kerf_count(stickData)
    kerfLength = _total_kerf_length(kerfCount, bladeKerf)
    partsLength = _total_parts_length(stickData)
    return abs(kerfLength + partsLength)

############ Private Helper Functions ############
def _inchToDecimalInch(measurement):
    return float(measurement) * 1000

def _orderList(list):
    return sorted(list)

def _total_parts_length(list):
    return sum(list)

def _total_kerf_count(list):
    return (len(list) - 1)

def _total_kerf_length(KerfCount, KerfWidth):
    return (KerfCount * KerfWidth)

def _ensure_negative(measurement):
    return (abs(measurement) * -1)

def _drop_first(list):
    return list[1:]