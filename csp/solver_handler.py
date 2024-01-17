from stock_cutter_1d import solveCut
from alns_stock_cutter import alnsSolver
from tkinter import messagebox

#TODO: Add a hybrid solver that uses ALNS to generate a good initial solution and then uses OR-Tools to optimize it.
class CuttingParameters:
    """A class used to represent and interact with the cutting parameters for the stick packing problem.

    :ivar int or float stock_length: The length of the stock material.
    :ivar int or float blade_width: The width of the blade.
    :ivar int or float dead_zone: The dead zone of the blade.
    :ivar list cut_lengths: The lengths of the cuts.
    :ivar list cut_quantities: The quantities of the cuts.
    :ivar int scale_factor: Scale factor used to convert data to decimal integer.
    :ivar int jobNumber: The job number.
    :ivar str dirPath: The directory path.
    :ivar str solver: The solver type.
    :ivar str author: The author of the solution.
    :ivar int staringProgramNumber: The starting program number.
    :ivar str fileName: The name of the Excel files.
    :ivar bool debug: The debug status.
    :ivar list of list of int or float solution: The solution to the stick packing problem.

    Dependencies:
        - solveCut from stock_cutter_1d module
        - alnsSolver from alns_stock_cutter module
        - messagebox from tkinter module
    """    
    def __init__(self, stock_length, blade_width, dead_zone, cut_lengths, cut_quantities):
        self.stock_length = stock_length
        self.blade_width = blade_width
        self.dead_zone = dead_zone
        self.cut_lengths = cut_lengths
        self.cut_quantities = cut_quantities
        self.scale_factor = None
        self.jobNumber = None
        self.dirPath = None
        self.solver = None
        self.author = None
        self.staringProgramNumber = None
        self.debug = False
        self.solution = None
        self.fileName = None

    def getStockLength(self):
        """Get the stock length for the cuttingParameters object.

        :return: Stock Length
        :rtype: int
        """        
        return self.stock_length

    def getBladeWidth(self):
        """Get the blade width for the cuttingParameters object.

        :return: Blade Width
        :rtype: int
        """        
        return self.blade_width

    def getStaringProgramNumber(self):
        """Get the starting program number for the cuttingParameters object.

        :return: Starting Program Number if set, else 1
        :rtype: int
        """
        if self.staringProgramNumber == None:
            return 1
        return self.staringProgramNumber

    def setStaringProgramNumber(self, startingProgramNumber):
        """Set the starting program number for the cuttingParameters object.

        :param int startingProgramNumber: Starting Program Number
        """
        self.staringProgramNumber = startingProgramNumber

    def getFileName(self):
        """Get the file name for the cuttingParameters object.

        :return: File Name
        :rtype: string
        """        
        return self.fileName
    
    def setFileName(self, fileName):
        """Set the file name for the cuttingParameters object.

        :param string fileName: File Name
        """        
        self.fileName = fileName
    
    def getScaleFactor(self):
        """Get the scale factor for the cuttingParameters object.

        :return: Scale Factor
        :rtype: int
        """        
        return self.scale_factor
    
    def setScaleFactor(self, scale_factor):
        """Set the scale factor for the cuttingParameters object.

        :param int scale_factor: Scale Factor
        """        
        self.scale_factor = scale_factor

    def getJobNumber(self):
        """Get the job number for the cuttingParameters object.

        :return: Job Number
        :rtype: int
        """        
        return self.jobNumber
    
    def setJobNumber(self, jobNumber):
        """Set the job number for the cuttingParameters object.

        :param int jobNumber: Job Number
        """        
        self.jobNumber = jobNumber

    def getDirPath(self):
        """Get the directory path for the cuttingParameters object.

        :return: Directory Path
        :rtype: String
        """        
        return self.dirPath
    
    def setDirPath(self, dirPath):
        """Sets the directory path for the cuttingParameters object.

        :param string dirPath: Directory Path
        """         
        self.dirPath = dirPath

    def getSolver(self):
        """Get the solver type for the cuttingParameters object.

        :return: Solver Type
        :rtype: string
        """       
        return self.solver
    
    def setSolver(self, solver):
        """Set the solver type for the cuttingParameters object.

        :param string solver: Solver Type
        """        
        self.solver = solver

    def getAuthor(self):
        """Get the author of the cuttingParameters object.

        :return: Author Name
        :rtype: string
        """        
        return self.author
    
    def setAuthor(self, author):
        """Sets the author for the cuttingParameters object.

        :param string author: Author Name
        """        
        self.author = author

    def getDebug(self):
        """Get the Debug status of the cuttingParameters object.

        :return: Debug Status
        :rtype: bool
        """        
        return self.debug
    
    def setDebug(self, debug):
        """Set the debug status of the cuttingParameters object.

        :param bool debug: Debug Status
        """        
        self.debug = debug

    def buildSolution(self):
        """Builds the solution for cutting parameters object.

        :raises ValueError: If invalid numeric values are entered.
        """        
        try:
            stock_length, zipped_data, blade_width = self._solverPreProcess()
            if self.solver == "OR-Tools":
                solution = _solveORTools(zipped_data, stock_length)
                solution = _ortoolsPostProcessor(solution, blade_width, self.scale_factor)
            elif self.solver == "ALNS":
                solution = _solveALNS(zipped_data, stock_length)
                solution = _alnsPostProcessor(solution, blade_width, self.scale_factor)
            self.solution = solution
        except ValueError as handler:
            print(f"Error: {handler}")
            messagebox.showerror("Error", "Please enter valid numeric values.")

    def getSolution(self):
        """Get the solution of the CuttingParameters object.

        :return: Solution
        :rtype: List of lists of floats
        """        
        return self.solution

    def print_solution(self):
        """Prints the solution of the stick packing problem.

        Example Output:
            - Stick 1: [1, 0, 1, 0, 1], Usage: 60.00%
            - Blade Width: 10, Dead Zone: 2
        """
        for idx, stick in enumerate(self.solution, start=1):
            usage = sum(stick) / (self.stock_length / self.scale_factor) * 100
            print(f"Stick {idx}: {stick}, Usage: {usage:.2f}%")
            print(f"Blade Width: {self.blade_width}, Dead Zone: {self.dead_zone}")

    def _solverPreProcess(self):
        stock_length = _scaleMeasurement(self.stock_length, self.scale_factor)
        blade_width = _scaleMeasurement(self.blade_width, self.scale_factor)
        cut_lengths = [_scaleMeasurement(length, self.scale_factor) for length in self.cut_lengths]
        zipped_data = _zipCutData(cut_lengths, self.cut_quantities)
        zipped_data = _addBladeKerf(zipped_data, blade_width)
        return stock_length, zipped_data, blade_width

def _solveORTools(zipped_data, stock_length):
    zipped_data = [[quantity, length] for length, quantity in zipped_data]
    return solveCut(zipped_data, stock_length, output_json=False, large_model=True, greedy_model=False, iterAccuracy=500)

def _solveALNS(zipped_data, stock_length):
    zipped_data = _flattenCutData(zipped_data)
    return alnsSolver(stock_length, zipped_data, iterations=1000, seed=1234)

def _ortoolsPostProcessor(solution, blade_width, scale_factor):
    return [[(length - blade_width) / scale_factor for length in stick[1]] for stick in solution]

def _alnsPostProcessor(solution, blade_width, scale_factor):
    return [[(length - blade_width) / scale_factor for length in assignments] for assignments in solution]

def _deScaleMeasurement(measurement, scaleFactor):
    return int(float(measurement) / scaleFactor)

def _scaleMeasurement(measurement, scaleFactor):
    return int(float(measurement) * scaleFactor)

def _zipCutData(cut_lengths, cut_quantities):
    return sorted(zip(cut_lengths, cut_quantities), key=lambda pair: pair[0], reverse=True)

def _flattenCutData(cutData):   
    return [length for (length, quantity) in cutData for _ in range(quantity)]

def _addBladeKerf(cutData, bladeKerf):
    return [[length + bladeKerf, quantity] for length, quantity in cutData]
