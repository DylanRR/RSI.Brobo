import xlsxwriter
import os
import re

class buildXFile:
    """
    Class for building an Excel file with specific data.

    Args:
        stickData (list): List of measurements.
        programNumber (int): Program number.
        jobNumber (str): Job number.
        dirPath (str): Directory path.
        fileName (str, optional): File name. Defaults to None.

    Usage:
        xlsFile = buildXFile(stickData, programNumber, jobNumber, dirPath, fileName) #fileName is optional
        xlsFile.setAuthor('NAME') # Optional
        xlsFile.buildSheet()
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
        """
        Builds the sheet by calling various methods to construct the header, author, program number, and cuts.
        Finally, it saves the file.
        
        Args:
            self: The current instance of the class.
        """
        self._buildHeader()
        self._buildAuthor()
        self._buildProgramNumber()
        self._buildCuts()
        self._saveFile()

    def setAuthor(self, author):
        """
        Set the author of the file.

        Args:
            self: The current instance of the class.
            author (str): The name of the author.
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
        Otherwise, it processes the file name using the `_processFileName` method and checks if it contains any invalid characters.
        If invalid characters are found, it prints a message and automatically generates a new file name.
        Finally, it checks if a file with the same name already exists in the save path. If it does, it prints a message and generates a new file name.

        Args:
            self: The current instance of the class.
        """
        self._buildJobNumPath()

        if not self.fileName:
            self._autoGenerateFileName()
        else:
            self._processFileName()
            if re.search(r'[\\/*?:"<>|]', self.fileName):
                print("Invalid characters in file name.")
                self._autoGenerateFileName()
            else:
                if os.path.exists(self.savePath + self.fileName):
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
        if not os.path.exists(self.savePath + '/' + self.jobNum):
            os.makedirs(self.savePath + '/' + self.jobNum)
        self.savePath = self.savePath + '/' + self.jobNum + '/'
    
    def _processFileName(self):
        """
        Process the file name by appending '.xlsx' if it doesn't already end with it.

        Args:
            self: The current instance of the class.
        """
        if not self.fileName.lower().endswith('.xlsx'):
            self.fileName += '.xlsx'
        
    def _autoGenerateFileName(self):
        """
        Auto-generates a unique file name for the Excel file based on the program number and the number of stick data parts.

        Args:
            self: The current instance of the class.
        """
        print("Auto Generating File Name........................")
        files = [f for f in os.listdir(self.savePath) if f.endswith('.xlsx')]
        count = len(files) + 1

        while True:
            count_str = str(count).zfill(3)
            file_name = count_str + '-Program_Num_'+ str(self.programNumber) +'_'+ str(len(self.stickData)) + '_parts.xlsx'
            print(f"File Name: {file_name}")
            if not os.path.exists(self.savePath + file_name):
                break
            count += 1
        self.fileName = file_name