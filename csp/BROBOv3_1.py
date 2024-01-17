from asyncio.windows_events import NULL
import tkinter as tk
from solver_handler import CuttingParameters
from brobo_preprocessor import buildBroboProgram

class CutOptimizerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Cut Optimizer")

        self.stock_length = tk.StringVar()
        self.blade_width = tk.StringVar()
        self.dead_zone = tk.StringVar()
        self.scale_factor = 100

        self.cut_lengths = []
        self.cut_quantities = []

        self.create_widgets()

    def create_widgets(self):
        self.root.geometry("270x400")  # Set the initial window size

        tk.Label(self.root, text="Stock Length:").pack()
        tk.Entry(self.root, textvariable=self.stock_length).pack()

        tk.Label(self.root, text="Blade Width:").pack()
        tk.Entry(self.root, textvariable=self.blade_width).pack()

        tk.Label(self.root, text="Dead Zone:").pack()
        tk.Entry(self.root, textvariable=self.dead_zone).pack()

        self.add_cut_button = tk.Button(self.root, text="Add Cut", command=self.add_cut)
        self.add_cut_button.pack()
        self.add_cut_button.config(state="disabled")  # Disable the button initially

        self.optimize_button = tk.Button(self.root, text="Optimize Cuts", command=self.UtestOptimize)
        self.optimize_button.pack()

        self.cut_canvas = tk.Canvas(self.root)
        self.cut_frame = tk.Frame(self.cut_canvas)

        self.cut_scrollbar = tk.Scrollbar(self.root, orient="vertical", command=self.cut_canvas.yview)
        self.cut_canvas.configure(yscrollcommand=self.cut_scrollbar.set)

        self.cut_scrollbar.pack(side="right", fill="y")
        self.cut_canvas.pack(side="left", fill="both", expand=True)
        self.cut_canvas.create_window((0, 0), window=self.cut_frame, anchor="nw")

        self.cut_frame.bind("<Configure>", self.on_frame_configure)

        tk.Label(self.cut_frame, text="Cut Length").grid(row=0, column=0)
        tk.Label(self.cut_frame, text="Quantity").grid(row=0, column=1)

        # Bind a function to be called whenever the entry fields are updated
        self.stock_length.trace_add("write", self.check_initial_params)
        self.blade_width.trace_add("write", self.check_initial_params)
        self.dead_zone.trace_add("write", self.check_initial_params)

    def on_frame_configure(self, event):
        self.cut_canvas.configure(scrollregion=self.cut_canvas.bbox("all"))

    def check_initial_params(self, *args):
        # Check if all three initial parameters are provided
        if self.stock_length.get() and self.blade_width.get() and self.dead_zone.get():
            self.add_cut_button.config(state="normal")  # Enable the button
        else:
            self.add_cut_button.config(state="disabled")  # Disable the button

    def add_cut(self):
        cut_length = tk.StringVar()
        cut_quantity = tk.StringVar()

        entry_cut_length = tk.Entry(self.cut_frame, textvariable=cut_length)
        entry_cut_quantity = tk.Entry(self.cut_frame, textvariable=cut_quantity)

        entry_cut_length.grid(row=len(self.cut_lengths) + 1, column=0)
        entry_cut_quantity.grid(row=len(self.cut_quantities) + 1, column=1)

        self.cut_lengths.append(cut_length)
        self.cut_quantities.append(cut_quantity)
    
    def getStockLength(self):
        return self.stock_length.get()
    
    def getBladeWidth(self):
        return self.blade_width.get()
    
    def getDeadZone(self):
        return self.dead_zone.get()
    
    def getCutLengths(self):
        return [cut_length.get() for cut_length in self.cut_lengths]
    
    def getCutQuantities(self):
        return [int(cut_quantity.get()) for cut_quantity in self.cut_quantities]
    
    def getScaleFactor(self):
        #TODO: return self.scale_factor.get()
        return 100
    
    def getSolver(self):
        #TODO: return self.solver.get()
        return "ALNS"
    
    def getAuthor(self):
        #TODO: return self.author.get()
        return "Dylan"
    
    def getJobNumber(self):
        #TODO: return self.jobNumber.get()
        return 1212
    
    def getDirPath(self):
        #TODO: return self.dirPath.get()
        return "U:/Git Development/rsi.Brobo"

    def getDebug(self):
        #TODO: return self.debug.get()
        return True
    
    

    def uTest(self):
        stock_length = 50
        blade_width = 1
        dead_zone = 5
        cut_lengths = [5, 10, 15, 20]
        cut_quantities = [1, 1, 1, 6]
        return stock_length, blade_width, dead_zone, cut_lengths, cut_quantities
        # inital Demand should be ~53000~

    def UtestOptimize(self):
        stock_length, blade_width, dead_zone, cut_lengths, cut_quantities = self.uTest()
        print(f"Stock Length: {stock_length}\nBlade Width: {blade_width}\nDead Zone: {dead_zone}\nCut Lengths: {cut_lengths}\nCut Quantities: {cut_quantities}")
        newCut = CuttingParameters(stock_length, blade_width, dead_zone, cut_lengths, cut_quantities)
        newCut.setSolver('ALNS')
        newCut.setScaleFactor(100)
        newCut.setJobNumber(1212)
        newCut.setDirPath("U:/Git Development/rsi.Brobo")
        newCut.setAuthor("Dylan")
        newCut.setStaringProgramNumber(4)
        newCut.setFileName("test")
        newCut.setDebug(True)
        newCut.buildSolution()
        newCut.print_solution()
        buildBroboProgram(newCut)

    def optimize(self):
        newCut = CuttingParameters(self.getStockLength(), self.getBladeWidth(), self.getDeadZone(), self.getCutLengths(), self.getCutQuantities())
        newCut.setSolver(self.getSolver())
        newCut.setScaleFactor(self.getScaleFactor())
        newCut.setJobNumber(self.getJobNumber())
        newCut.setDirPath(self.getDirPath())
        newCut.setAuthor(self.getAuthor())
        newCut.setDebug(self.getDebug())
        newCut.buildSolution()
        solution = newCut.getSolution()
        print (solution)
        #buildBroboProgram(solution, newCut.getBladeWidth(), newCut.getJobNumber(), newCut.getDirPath())

if __name__ == "__main__":
    root = tk.Tk()
    app = CutOptimizerApp(root)
    root.mainloop()
