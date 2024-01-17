import xlsxwriter
import os
import re

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
        self.workbook = xlsxwriter.Workbook(self.savePath + self.fileName)
        self.worksheet = self.workbook.add_worksheet()

    def buildSheet(self):
        """This method fully builds the Excel file.
        """                
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
        self.author = author

    ############ Private Methods ############
    def _buildHeader(self):
        """
        Builds the header row in the worksheet.

        The header row contains the column names for the different attributes.

        Args:
            self: The instance of the class.
        """
        self.worksheet.write('A1', 'axis')
        self.worksheet.write('B1', 'number')
        self.worksheet.write('C1', 'demand')
        self.worksheet.write('D1', 'quantity')
        self.worksheet.write('E1', 'mode')
        self.worksheet.write('F1', 'output')

    def _buildAuthor(self):
        """
        Builds the author information in the worksheet.
        If the author is provided, it writes the author's name in cell G1 of the worksheet.

        Args:
            self: The current instance of the class.
        """
        if self.author:
            self.worksheet.write('G1', ';Program Author: ' + self.author)

    def _buildProgramNumber(self):
        """
        Builds the program number in the worksheet.
        This method writes the program number to cells A2 and B2 in the worksheet.

        Args:
            self: The current instance of the class.
        """
        self.worksheet.write('A2', 'Pno')
        self.worksheet.write('B2', self.programNumber)
    
    def _buildCuts(self):
        """
        Builds the cuts in the worksheet based on the stickData.

        Args:
            self: The current instance of the class.
        """
        tempLineNum = 3
        for measurement in self.stickData:
            self.worksheet.write('A' + str(tempLineNum), 0)
            self.worksheet.write('B' + str(tempLineNum), tempLineNum - 2)
            self.worksheet.write('C' + str(tempLineNum), measurement)
            self.worksheet.write('D' + str(tempLineNum), 1)
            modeVal = 0 if tempLineNum == 3 else 1
            self.worksheet.write('E' + str(tempLineNum), modeVal)
            self.worksheet.write('F' + str(tempLineNum), 1)
            tempLineNum += 1
        self._buildFinalLine(tempLineNum)

    def _saveFile(self):
        """
        Closes the workbook and saves the file.
        """
        self.workbook.close()

    def _buildFinalLine(self, lineNum):
        """
        Builds the final line in the worksheet.

        Args:
            self: The current instance of the class.
            lineNum (int): The line number.
        """
        self.worksheet.write('A' + str(lineNum), 0)
        self.worksheet.write('B' + str(lineNum), lineNum)
        self.worksheet.write('C' + str(lineNum), 0)
        self.worksheet.write('D' + str(lineNum), 0)
        self.worksheet.write('E' + str(lineNum), 0)
        self.worksheet.write('F' + str(lineNum), 1)


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

        if not self.fileName.lower().endswith('.xlsx'):
            self.fileName += '.xlsx'

    def _autoGenerateFileName(self):
        """
        Auto-generates a unique file name for the Excel file based on the program number and the number of stick data parts.

        Args:
            self: The current instance of the class.
        """
        print("Auto Generating File Name........................")
        
        count = 0
        highest_number = 0
        files = [f for f in os.listdir(self.savePath) if f.endswith('.xlsx')]
        for file in files:
            match = re.match(r'^(\d{3})', file)
            if match:
                number = int(match.group(1))
                if number > highest_number:
                    highest_number = number

        count = highest_number + 1 if highest_number > len(files) else len(files) + 1

        while True:
            count_str = str(count).zfill(3)
            file_name = count_str + '-Program_Num_'+ str(self.programNumber) +'_'+ str(len(self.stickData)) + '_parts.xlsx'
            print(f"File Name: {file_name}")
            if not os.path.exists(self.savePath + file_name):
                break
            count += 1
        self.fileName = file_name