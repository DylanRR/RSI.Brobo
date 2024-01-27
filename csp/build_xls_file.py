import xlwt
import os
import re

DEBUG = True
class buildXFile:
    """
    A class used to build xls files in a format that a BROBO Semi-Automatic cold saw can interpret.

    :ivar list stickData: Individual stick data.
    :ivar int programNumber: The program number.
    :ivar int jobNum: The job number.
    :ivar str savePath: The save path.
    :ivar str fileName: (Optional) The file name.
    :ivar str author: The author of the solution.

    Dependencies:
        - xlsxwriter
        - os
        - re
    """
    def __init__(self, stickData, programNumber,  jobNumber, dirPath,  fileName = None):
        self.stickData = stickData
        self.savePath = dirPath
        self.jobNum = jobNumber
        self.programNumber = programNumber
        self.author = 'Auto Generated'
        self.fileName = fileName
        self._checkValidSaveLocation()
        self.workbook = xlwt.Workbook()
        self.worksheet = self.workbook.add_sheet('Sheet 1')

    def buildSheet(self):
        """This method fully builds the Excel file.
        """                
        if DEBUG:
            print("Building Sheet........................")

        self._buildHeader()
        self._buildAuthor()
        self._buildProgramNumber()
        self._buildCuts()
        self._saveFile()

    def setAuthor(self, author):
        """Used to set the author of the solution.

        :param author: Author Name
        :type author: String
        """
        if DEBUG:
            print("Setting Author........................")

        self.author = author

    ############ Private Methods ############
    def _buildHeader(self):
        """
        Builds the header row in the worksheet.

        The header row contains the column names for the different attributes.

        Args:
            self: The instance of the class.
        """
        if DEBUG:
            print("Building Header........................")

        self.worksheet.write(0, 0, 'axis')     #A1 Write
        self.worksheet.write(0, 1, 'number')   #B1 Write
        self.worksheet.write(0, 2, 'demand')   #C1 Write
        self.worksheet.write(0, 3, 'quantity') #D1 Write
        self.worksheet.write(0, 4, 'mode')     #E1 Write
        self.worksheet.write(0, 5, 'output')   #F1 Write

    def _buildAuthor(self):
        """
        Builds the author information in the worksheet.
        If the author is provided, it writes the author's name in cell G1 of the worksheet.

        Args:
            self: The current instance of the class.
        """
        if DEBUG:
            print("Building Author........................")

        if self.author:
            self.worksheet.write(0, 6, ';Program Author: ' + self.author) #G1 Write

    def _buildProgramNumber(self):
        """
        Builds the program number in the worksheet.
        This method writes the program number to cells A2 and B2 in the worksheet.

        Args:
            self: The current instance of the class.
        """
        if DEBUG:
            print("Building Program Number........................")
        self.worksheet.write(1, 0, 'Pno') #A2 Write
        self.worksheet.write(1, 1, self.programNumber)  #B2 Write
    
    def _buildCuts(self):
        """
        Builds the cuts in the worksheet based on the stickData.

        Args:
            self: The current instance of the class.
        """
        if DEBUG:
            print("Building Cuts........................")

        tempLineNum = 3
        for measurement in self.stickData:
            self.worksheet.write(tempLineNum - 1, 0, 0) #A3 Write
            self.worksheet.write(tempLineNum - 1, 1, tempLineNum - 2) #B3 Write
            self.worksheet.write(tempLineNum - 1, 2, measurement) #C3 Write
            self.worksheet.write(tempLineNum - 1, 3, 1) #D3 Write
            modeVal = 0 if tempLineNum == 3 else 1
            self.worksheet.write(tempLineNum - 1, 4, modeVal) #E3 Write
            self.worksheet.write(tempLineNum - 1, 5, 1) #F3 Write
            tempLineNum += 1
        self._buildFinalLine(tempLineNum)

    def _saveFile(self):
        """
        Closes the workbook and saves the file.
        """
        if DEBUG:
            print("Saving File........................")

        self.workbook.save(self.savePath + self.fileName)
        # self.workbook.close() #Legacy Code from xlsxwriter library

    def _buildFinalLine(self, lineNum):
        """
        Builds the final line in the worksheet.

        Args:
            self: The current instance of the class.
            lineNum (int): The line number.
        """
        self.worksheet.write(lineNum - 1, 0, 0)       #A3 Write
        self.worksheet.write(lineNum - 1, 1, lineNum) #B3 Write
        self.worksheet.write(lineNum - 1, 2, 0)       #C3 Write
        self.worksheet.write(lineNum - 1, 3, 0)       #D3 Write
        self.worksheet.write(lineNum - 1, 4, 0)       #E3 Write
        self.worksheet.write(lineNum - 1, 5, 1)       #F3 Write

    def _checkValidSaveLocation(self):
        """
        Checks if the save location is valid and performs necessary actions.

        This method first builds the job number path using the `_buildJobNumPath` method.
        If the `fileName` attribute is empty, it automatically generates a file name using the `_autoGenerateFileName` method.
        Otherwise, it processes the file name using the `_processFileName` method.
        Finally, it checks if a file with the same name already exists in the save path. If it does, it generates a new file name.

        Args:
            self: The current instance of the class.
        """
        self._buildJobNumPath()

        if not self.fileName:
            self._autoGenerateFileName()
        else:
            self._processFileName()
            if os.path.exists(self.savePath + self.fileName):   #Only gets triggered in rare cases due to the _processFileName method doing initial checks (May be removed in the future)
                print("File already exists.")
                self._autoGenerateFileName()
            
    def _buildJobNumPath(self):
        """
        Creates a directory path based on the job number and saves it to the `savePath` attribute.

        If the directory path does not exist, it will be created using the `savePath` and `jobNum` attributes.
        The `savePath` attribute will be updated to include the job number directory.

        Args:
            self: The current instance of the class.
        """
        if DEBUG:
            print("Building Job Number Path........................")

        if not os.path.exists(self.savePath + '/' + str(self.jobNum)):
            os.makedirs(self.savePath + '/' + str(self.jobNum))
        self.savePath = self.savePath + '/' + str(self.jobNum) + '/'
    
    def _processFileName(self):
        """
        Process the file name by first checking if the name contains invalid characters. If it does it will automatically generate a new file name.
        If the file name does not start with three digits, it will automatically add a three digit number to the beginning of the file name.
        If the file name does not end with '.xlsx', it will automatically add '.xlsx' to the end of the file name.

        Args:
            self: The current instance of the class.
        """

        if DEBUG:
            print("Processing File Name........................")

        if re.search(r'[\\/*?:"<>|]', self.fileName):
                print("Invalid characters in file name.")
                self._autoGenerateFileName()
                return
        
        if not re.match(r'^\d{3}', self.fileName):
            print("Filename does not start with three digits.")
            files = [f for f in os.listdir(self.savePath) if f.endswith('.xlsx')]

            while True:
                count = len(files) + 1
                count_str = str(count).zfill(3)
                fileName = count_str + '-' + self.fileName
                if not os.path.exists(self.savePath + fileName):
                    break
                count += 1
            self.fileName = fileName

        if not self.fileName.lower().endswith('.xls'):
            self.fileName += '.xls'

    def _autoGenerateFileName(self):
        """
        Auto-generates a unique file name for the Excel file based on the program number and the number of stick data parts.

        Args:
            self: The current instance of the class.
        """
        if DEBUG:
            print("Auto Generating File Name........................")
        
        count = 0
        highest_number = 0
        files = [f for f in os.listdir(self.savePath) if f.endswith('.xls')]
        for file in files:
            match = re.match(r'^(\d{3})', file)
            if match:
                number = int(match.group(1))
                if number > highest_number:
                    highest_number = number

        count = highest_number + 1 if highest_number > len(files) else len(files) + 1

        while True:
            count_str = str(count).zfill(3)
            file_name = count_str + '-Program_Num_'+ str(self.programNumber) +'_'+ str(len(self.stickData)) + '_parts.xls'
            print(f"File Name: {file_name}")
            if not os.path.exists(self.savePath + file_name):
                break
            count += 1
        self.fileName = file_name