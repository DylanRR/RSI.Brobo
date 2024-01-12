import xlsxwriter
import os
import re

class buildXFile:
    def __init__(self, stickData, programNumber,  jobNumber, dirPath,  fileName = None):
        self.stickData = stickData
        self.savePath = dirPath
        self.jobNum = jobNumber
        self.programNumber = programNumber
        self.author = 'Auto Generated'
        self.fileName = fileName
        self._checkValidSaveLocation()
        #print(f"save path: {self.savePath + self.fileName}")
        self.workbook = xlsxwriter.Workbook(self.savePath + self.fileName)
        self.worksheet = self.workbook.add_worksheet()

    def buildSheet(self):
        self._buildHeader()
        self._buildAuthor()
        self._buildProgramNumber()
        self._buildCuts()
        self._saveFile()

    def setAuthor(self, author):
        self.author = author

    def setFileName(self, fileName):
        self.fileName = fileName

    def _buildHeader(self):
        self.worksheet.write('A1', 'axis')
        self.worksheet.write('B1', 'number')
        self.worksheet.write('C1', 'demand')
        self.worksheet.write('D1', 'quantity')
        self.worksheet.write('E1', 'mode')
        self.worksheet.write('F1', 'output')

    def _buildProgramNumber(self):
        self.worksheet.write('A2', 'Pno')
        self.worksheet.write('B2', self.programNumber)
    
    def _buildCuts(self):
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

    def _buildFinalLine(self, lineNum):
        self.worksheet.write('A' + str(lineNum), 0)
        self.worksheet.write('B' + str(lineNum), lineNum)
        self.worksheet.write('C' + str(lineNum), 0)
        self.worksheet.write('D' + str(lineNum), 0)
        self.worksheet.write('E' + str(lineNum), 0)
        self.worksheet.write('F' + str(lineNum), 1)

    def _buildAuthor(self):
        if self.author:
            self.worksheet.write('G1', ';Program Author: ' + self.author)


    def _checkValidSaveLocation(self):
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
        if not os.path.exists(self.savePath + '/' + self.jobNum):
            os.makedirs(self.savePath + '/' + self.jobNum)
        self.savePath = self.savePath + '/' + self.jobNum + '/'
    
    def _processFileName(self):
        if not self.fileName.lower().endswith('.xls'):
            self.fileName += '.xls'
        
    def _autoGenerateFileName(self):
        print("Auto Generating File Name........................")
        files = [f for f in os.listdir(self.savePath) if f.endswith('.xls')]
        count = len(files) + 1

        while True:
            count_str = str(count).zfill(3)
            file_name = count_str + '-Program_Num_'+ str(self.programNumber) +'_'+ str(len(self.stickData)) + '_parts.xls'
            print(f"File Name: {file_name}")
            if not os.path.exists(self.savePath + file_name):
                break
            count += 1
        self.fileName = file_name
        
    def _saveFile(self):
        self.workbook.close()