'''
Contributions:
            Original Author: Serge Kruk https://github.com/sgkruk/Apress-AI/blob/master/cutting_stock.py
            Updated V2: Emad Ehsan https://github.com/emadehsan/Apress-AI/blob/master/my-models/custom_cutting_stock.py
            Updated V3: Dylan Ragaishis https://github.com/DylanRR
'''
from ortools.linear_solver import pywraplp
from math import ceil


"""
    Create and return a new solver instance.

    Args:
        name (str): The name of the solver.
        integer (bool): Set to True for an integer programming solver, False for linear programming.

    Returns:
        pywraplp.Solver: A new solver instance.
    """
def newSolver(name,integer=False):
  return pywraplp.Solver(name, pywraplp.Solver.CBC_MIXED_INTEGER_PROGRAMMING if integer else pywraplp.Solver.GLOP_LINEAR_PROGRAMMING)


"""
    Extract solution values from a solver result, handling various data types.

    Args:
        x: The input object representing a solver result or value.

    Returns:
        Union[int, float, List[Union[int, float]]]: Extracted solution value(s).
"""
def SolVal(x):
  if type(x) is not list:
    return 0 if x is None \
      else x if isinstance(x,(int,float)) \
      else x.SolutionValue() if x.Integer() is False \
      else int(x.SolutionValue())
  elif type(x) is list:
    return [SolVal(e) for e in x]



"""
    Retrieve the objective value of a solver solution.

    Args:
        x: The solver solution object.

    Returns:
        float: The objective value of the solution.
    """
def ObjVal(x):
  return x.Objective().Value()


"""
    Solve the Cutting Stock Problem using a small model approach.

    This function solves the Cutting Stock Problem using a small model approach, where it creates integer variables
    for big stick usage and cuts. It then formulates and adds constraints to ensure demand fulfillment and
    maximum size limits. The objective is to minimize the total number of big sticks used. The function returns
    the status of the solver, the number of big sticks used, the consumed big sticks, the unused stick widths,
    and the wall time taken for solving.

    Args:
        demands (List[List[int]]): A list of demand quantities and widths for small sticks.
        parent_width (int, optional): Width of the parent stick. Defaults to 100.

    Returns:
        Tuple: A tuple containing the solver status, number of big sticks used, consumed big sticks,
        unused stick widths, and wall time taken.
"""
def solve_model(demands, parent_width=100):
  num_orders = len(demands)
  solver = newSolver('Cutting Stock', True)
  k,b  = bounds(demands, parent_width)
  
  # Create variables for big stick usage and cuts
  y = [ solver.IntVar(0, 1, f'y_{i}') for i in range(k[1]) ] 
  x = [[solver.IntVar(0, b[i], f'x_{i}_{j}') for j in range(k[1])] \
      for i in range(num_orders)]
  unused_widths = [ solver.NumVar(0, parent_width, f'w_{j}') \
      for j in range(k[1]) ] 
  nb = solver.IntVar(k[0], k[1], 'nb')

  # consntraint: demand fullfilment
  for i in range(num_orders):  
    solver.Add(sum(x[i][j] for j in range(k[1])) >= demands[i][0]) 

  # constraint: max size limit
  for j in range(k[1]):
    solver.Add(sum(demands[i][1]*x[i][j] for i in range(num_orders)) <= parent_width*y[j])
    solver.Add(parent_width*y[j] - sum(demands[i][1]*x[i][j] for i in range(num_orders)) == unused_widths[j])
    if j < k[1]-1:
      solver.Add(sum(x[i][j] for i in range(num_orders)) >= sum(x[i][j+1] for i in range(num_orders)))

  solver.Add(nb == solver.Sum(y[j] for j in range(k[1])))

  # Minimize total big sticks used
  Cost = solver.Sum((j+1)*y[j] for j in range(k[1]))
  solver.Minimize(Cost)

  status = solver.Solve()
  numSticksUsed = SolVal(nb)

  return status, numSticksUsed, sticks(numSticksUsed, SolVal(x), SolVal(unused_widths), demands), SolVal(unused_widths), solver.WallTime()



"""
    Calculate bounds for the Cutting Stock Problem.

    This function calculates the lower and upper bounds for the Cutting Stock Problem, which are used to guide
    the optimization process. It determines the minimum number of big sticks required (lower bound) and the maximum
    number of small sticks that can be consumed (upper bound) based on the given demands and parent stick width.

    Args:
        demands (List[List[int]]): A list of demand quantities and widths for small sticks.
        parent_width (int, optional): Width of the parent stick. Defaults to 100.

    Returns:
        Tuple[List[int], List[int]]: A tuple containing the lower bound (k) and the upper bound (b) for the problem.
 """
def bounds(demands, parent_width=100):
  num_orders = len(demands)
  b = []
  T = 0
  k = [0,1]
  TT = 0

  for i in range(num_orders):
    quantity, width = demands[i][0], demands[i][1]
    b.append( min(quantity, int(round(parent_width / width))) )

    if T + quantity*width <= parent_width:
      T, TT = T + quantity*width, TT + quantity*width
    else:
      while quantity:
        if T + width <= parent_width:
          T, TT, quantity = T + width, TT + width, quantity-1
        else:
          k[1],T = k[1]+1, 0
  k[0] = int(round(TT/parent_width+0.5))

  print('k', k)
  print('b', b)
  return k, b



"""
    Generate the list of consumed big sticks based on the solution.

    This function generates a list of consumed big sticks based on the solution obtained from the optimization model.
    It calculates the width of each stick used and the corresponding small sticks that are cut from it.

    Args:
        nb (int): Number of big sticks used.
        x (List[List[int]]): Matrix representing the quantity of small sticks cut from each order for each big stick.
        w (List[int]): List of unused widths for each big stick.
        demands (List[List[int]]): A list of demand quantities and widths for small sticks.

    Returns:
        List[List[Union[int, List[int]]]]: A list of consumed big sticks, where each entry contains the width of
        the big stick and the corresponding small sticks cut from it.
 """
def sticks(nb, x, w, demands):
  consumed_big_sticks = []
  num_orders = len(x) 

  for j in range(len(x[0])):
    RR = [ abs(w[j])] + [ int(x[i][j])*[demands[i][1]] for i in range(num_orders) if x[i][j] > 0 ]
    consumed_big_sticks.append(RR)

  return consumed_big_sticks



"""
    Solve the large-scale cutting stock model using Dantzig-Wolfe decomposition.

    This function solves the large-scale cutting stock optimization problem using Dantzig-Wolfe decomposition.
    It iteratively improves the solution by optimizing patterns and pattern usage. The goal is to minimize the
    number of big sticks used while satisfying demand quantities for small sticks.

    Args:
        demands (List[List[int]]): A list of demand quantities and widths for small sticks.
        parent_width (int, optional): The width of the parent stick. Defaults to 100.
        iterAccuracy (int, optional): The number of column generations to iterate through. Defaults to 20.

    Returns:
        tuple: A tuple containing the solver status, optimized patterns, pattern usage (y),
               and the sticks cut using the optimized patterns.
 """
def solve_large_model(demands, parent_width=100, iterAccuracy=20):
  num_orders = len(demands)
  iter = 0
  patterns = get_initial_patterns(demands)
  quantities = [demands[i][0] for i in range(num_orders)]
  print('quantities', quantities)

  while iter < iterAccuracy:
    status, y, l = solve_master(patterns, quantities, parent_width=parent_width)
    iter += 1

    widths = [demands[i][1] for i in range(num_orders)]
    new_pattern, objectiveValue = get_new_pattern(l, widths, parent_width=parent_width)

    for i in range(num_orders):
      patterns[i].append(new_pattern[i])

  status, y, l = solve_master(patterns, quantities, parent_width=parent_width, integer=True)  

  return status, patterns, y, sticks_patterns(patterns, y, demands, parent_width=parent_width)



"""
    Solve the master problem for the cutting stock model.

    This function defines and solves the master problem, which is used in the context of the cutting stock
    optimization problem. The master problem aims to find the optimal combination of pattern usage that minimizes
    the number of big sticks used.

    Args:
        patterns (List[List[int]]): A list of patterns where each pattern is a list of integers representing the usage of each small stick for that pattern.
        quantities (List[int]): A list of integers representing the demand quantities for each pattern.
        parent_width (int, optional): The width of the parent stick. Defaults to 100.
        integer (bool, optional): If True, the solver uses integer programming, otherwise linear programming. (Defaults to False)

    Returns:
        tuple: A tuple containing the status of the solver, a list of optimized pattern usage (y), 
               and a list of dual values (l) associated with the constraints.
 """
def solve_master(patterns, quantities, parent_width=100, integer=False):
  title = 'Cutting stock master problem'
  num_patterns = len(patterns)
  n = len(patterns[0])
  constraints = []

  solver = newSolver(title, integer)
  
  y = [ solver.IntVar(0, 1000, '') for j in range(n) ]
  Cost = sum(y[j] for j in range(n)) 
  solver.Minimize(Cost)

  for i in range(num_patterns):
    constraints.append(solver.Add( sum(patterns[i][j]*y[j] for j in range(n)) >= quantities[i]) ) 

  status = solver.Solve()
  y = [int(ceil(e.SolutionValue())) for e in y]

  l =  [0 if integer else constraints[i].DualValue() for i in range(num_patterns)]
  toreturn = status, y, l
  return toreturn



"""
    Solve the sub-problem to find a new cutting pattern.

    This function solves a sub-problem to find a new cutting pattern based on given dual values (l) and
    unused stick widths (w). It creates integer variables for the new pattern and maximizes the objective,
    which is the sum of dual values multiplied by the new pattern variables. Constraints are added to ensure
    the total width of the new pattern does not exceed the parent stick width. The function returns the new
    pattern and the objective value of the solver.

    Args:
        l (List[float]): List of dual values corresponding to constraints.
        w (List[float]): List of unused stick widths for each constraint.
        parent_width (int, optional): Width of the parent stick. Defaults to 100.

    Returns:
        Tuple: A tuple containing the new cutting pattern and the objective value of the solver.
    """
def get_new_pattern(l, w, parent_width=100):
  solver = newSolver('Cutting stock sub-problem', True)
  n = len(l)
  new_pattern = [ solver.IntVar(0, parent_width, '') for i in range(n) ]

  Cost = sum( l[i] * new_pattern[i] for i in range(n))
  solver.Maximize(Cost)

  solver.Add( sum( w[i] * new_pattern[i] for i in range(n)) <= parent_width ) 

  status = solver.Solve()
  return SolVal(new_pattern), ObjVal(solver)



"""
    Generate initial cutting patterns.

    This function generates the initial cutting patterns based on the number of orders (demands).
    Each pattern corresponds to one order, and only the respective order's stick is used in the pattern.
    All other sticks are set to 0. These initial patterns are used to start the optimization process.

    Args:
        demands (List[List[int]]): List of order quantities and widths.

    Returns:
        List[List[int]]: List of initial cutting patterns, where each pattern has a single stick used.
    """
def get_initial_patterns(demands):
  num_orders = len(demands)
  return [[0 if j != i else 1 for j in range(num_orders)]\
          for i in range(num_orders)]



"""
    Generate detailed cutting patterns based on the optimized solution.

    This function generates detailed cutting patterns based on the optimized solution obtained from the large model.
    It takes the optimized patterns, the number of big sticks used, the original order demands, and the parent stick width.
    It creates a detailed breakdown of how each big stick is used, including the widths of small sticks cut from it.
    
    Args:
        patterns (List[List[int]]): List of optimized cutting patterns for each order.
        y (List[int]): List of the number of big sticks used for each pattern.
        demands (List[List[int]]): List of original order quantities and widths.
        parent_width (int, optional): Width of the parent stick. Defaults to 100.

    Returns:
        List[List[Union[int, List[int]]]]: Detailed breakdown of how each big stick is used, including unused width.
    """
def sticks_patterns(patterns, y, demands, parent_width=100):
  R, m, n = [], len(patterns), len(y)

  for j in range(n):
    for _ in range(y[j]):
      RR = []
      for i in range(m):
        if patterns[i][j] > 0:
          RR.extend( [demands[i][1]] * int(patterns[i][j]) )
      used_width = sum(RR)
      R.append([parent_width - used_width, RR])

  return R



"""
    Check if small stick widths are within acceptable limits.

    This function checks whether the widths of small sticks specified in the demands list
    are within acceptable limits, i.e., they do not exceed the width of the parent sticks.
    
    Args:
        demands (List[List[int]]): List of order quantities and widths.
        parent_width (int): Width of the parent stick.

    Returns:
        bool: True if all small stick widths are within limits, False otherwise.
"""
def checkWidths(demands, parent_width):
  for quantity, width in demands:
    if width > parent_width:
      print(f'Small stick width {width} is greater than parent sticks width {parent_width}. Exiting')
      return False
  return True


"""
    Solve the cutting stock problem using a greedy approach.

    Args:
        sorted_pairs (List[Tuple[int, int]]): List of tuples representing available cut lengths and their quantities.
        working_length (int): Remaining length of the current working stick.
        working_stick (List[int]): List of cut lengths on the current working stick.
        blade_width (int): Width of the blade that separates cut lengths.

    Returns:
        Tuple[List[int], List[Tuple[int, int]]] or None: A tuple containing the list of cut lengths on the working stick
        and the updated list of sorted cut pairs. If a solution is not possible, returns None.
"""


"""
    Solve the cutting stock problem using the specified approach.

    Args:
        cut_list (List[Tuple[int, int]]): List of tuples representing available cut lengths and their quantities.
        mtrl_length (int): Length of the material being cut.
        blade_width (int): Width of the blade that separates cut lengths.
        output_json (bool): If True, return the results in JSON format. If False, return a list.
        large_model (bool): If True, use the large cutting stock model. If False, use the small model.
        greedy_model (bool): If True, solve using a greedy approach. If False, use the specified model.

    Returns:
        List or str: Depending on the value of output_json, either a list of consumed sticks or a JSON string.
    """
def solveCut(cutData, stock_length, output_json=False, large_model=True, greedy_model=False, iterAccuracy=20):
    stock_length = [[1, stock_length]]
    solved = StockCutter1D(cutData, stock_length, output_json, large_model, iterAccuracy=iterAccuracy)
    return solved




"""
    Perform 1D stock cutting optimization using cutting stock algorithms.

    This function takes a list of child sticks and a list of parent sticks, and applies
    a 1D stock cutting optimization algorithm to find an efficient way to cut child sticks
    from parent sticks, minimizing wastage.

    Args:
        child_sticks (List[List[int]]): List of child stick quantities and widths.
        parent_sticks (List[List[int]]): List of parent stick quantities and widths.
        output_json (bool): If True, the output will be in JSON format, else in a list format.
        large_model (bool): If True, uses a large-scale optimization model, else uses a small model.

    Returns:
        List or str: If output_json is True, returns the output in JSON format, else as a list.

    Note:
        The function internally uses different algorithms based on the value of large_model:
        - If large_model is False, it uses a small-scale model for optimization.
        - If large_model is True, it uses a large-scale model for optimization.
"""
def StockCutter1D(child_sticks, parent_sticks, output_json=True, large_model=True, iterAccuracy=20):
  parent_width = parent_sticks[0][1]

  if not checkWidths(demands=child_sticks, parent_width=parent_width):
    return []

  print('child_sticks', child_sticks)
  print('parent_sticks', parent_sticks)

  if not large_model:
    print('Running Small Model...')
    status, numSticksUsed, consumed_big_sticks, unused_stick_widths, wall_time = \
              solve_model(demands=child_sticks, parent_width=parent_width)

    print('consumed_big_sticks before adjustment: ', consumed_big_sticks)
    new_consumed_big_sticks = []
    for big_stick in consumed_big_sticks:
      if len(big_stick) < 2:
        consumed_big_sticks.remove(big_stick)
        continue
      unused_width = big_stick[0]
      substicks = []
      for subitem in big_stick[1:]:
        if isinstance(subitem, list):
          substicks = substicks + subitem
        else:
          substicks.append(subitem)
      new_consumed_big_sticks.append([unused_width, substicks])
    print('consumed_big_sticks after adjustment: ', new_consumed_big_sticks)
    consumed_big_sticks = new_consumed_big_sticks
  
  else:
    print('Running Large Model...');
    status, A, y, consumed_big_sticks = solve_large_model(demands=child_sticks, parent_width=parent_width, iterAccuracy=iterAccuracy)

  numSticksUsed = len(consumed_big_sticks)

  STATUS_NAME = ['OPTIMAL',
    'FEASIBLE',
    'INFEASIBLE',
    'UNBOUNDED',
    'ABNORMAL',
    'NOT_SOLVED'
    ]

  output = {
      "statusName": STATUS_NAME[status],
      "numSolutions": '1',
      "numUniqueSolutions": '1',
      "numSticksUsed": numSticksUsed,
      "solutions": consumed_big_sticks
  }

  print('numSticksUsed', numSticksUsed)
  print('Status:', output['statusName'])
  print('Solutions found :', output['numSolutions'])
  print('Unique solutions: ', output['numUniqueSolutions'])

  if output_json:
    return json.dumps(output)        
  else:
    return consumed_big_sticks