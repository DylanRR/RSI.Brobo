from stock_cutter_1d import solveCut
from alns_stock_cutter import alnsSolver
from brobo_preprocessor import buildBroboProgram
from tkinter import messagebox

#TODO: Add a hybrid solver that uses ALNS to generate a good initial solution and then uses OR-Tools to optimize it.
class CuttingParameters:
    def __init__(self, stock_length, blade_width, dead_zone, cut_lengths, cut_quantities):
        self.stock_length = stock_length
        self.blade_width = blade_width
        self.dead_zone = dead_zone
        self.cut_lengths = cut_lengths
        self.cut_quantities = cut_quantities
        self.scale_factor = 100
        self.jobNumber = 0000
        self.dirPath = "C:/Users/Mothershipv2/Desktop/cnc.brobo/RSI.Brobo"
        self.solver = "ALNS"
        self.author = None
        self.debug = False
        self.solution = None

    def buildSolution(self):
        try:
            stock_length, blade_width, zipped_data = _solverPreProcess(self.stock_length, self.blade_width, self.dead_zone, self.cut_lengths, self.cut_quantities, self.scale_factor)
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
        return self.solution

    def buildBroboProgram(self):
        buildBroboProgram(self.solution, self.blade_width, self.jobNumber, self.dirPath)

    def print_solution(self):
        for idx, stick in enumerate(self.solution, start=1):
            usage = sum(stick) / (self.stock_length / self.scale_factor) * 100
            print(f"Stick {idx}: {stick}, Usage: {usage:.2f}%")
            print(f"Blade Width: {self.blade_width}, Dead Zone: {self.dead_zone}")

    def getJobNumber(self):
        return self.jobNumber
    
    def setJobNumber(self, jobNumber):
        self.jobNumber = jobNumber

    def getDirPath(self):
        return self.dirPath
    
    def setDirPath(self, dirPath):
        self.dirPath = dirPath

    def getSolver(self):
        return self.solver
    
    def setSolver(self, solver):
        self.solver = solver

    def getAuthor(self):
        return self.author
    
    def setAuthor(self, author):
        self.author = author

    def getDebug(self):
        return self.debug
    
    def setDebug(self, debug):
        self.debug = debug


def _solverPreProcess(stock_length, blade_width, dead_zone, cut_lengths, cut_quantities, scale_factor):
    stock_length = _scaleMeasurement(stock_length, scale_factor)
    blade_width = _scaleMeasurement(blade_width, scale_factor)
    dead_zone = _scaleMeasurement(dead_zone, scale_factor)
    stock_length = stock_length - dead_zone
    cut_lengths = [_scaleMeasurement(cut_length, scale_factor) for cut_length in cut_lengths]
    zipped_data = _zipCutData(cut_lengths, cut_quantities)
    zipped_data = _addBladeKerf(zipped_data, blade_width)
    return stock_length, blade_width, dead_zone, cut_lengths, zipped_data

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
