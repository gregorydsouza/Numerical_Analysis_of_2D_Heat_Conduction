import sympy
import xlsxwriter
import re
import pandas as pd

# Properties
k = 8.0 # Btu / hr*ft*degF
h = 96.0 # Btu / hr*ft*ft*degF
T_inf = 80.0 # degF
T_b = 200.0 # degF
dx = (1.0 / 8.0) / 12.0 # inch
dy = dx # inch

# k = sympy.symbols("k")
# h = sympy.symbols("h")
# T_inf = sympy.symbols(r"T_{\infty}")
# T_b = sympy.symbols("T_b")
# dx = sympy.symbols(r"\Delta{x}")
# dy = sympy.symbols(r"\Delta{y}")

# Create a class to generalize each node
class TNode:
	# Neighbors
	bn = None # Bottom neighbor
	tn = None # Top neighbor
	ln = None # Left neighbor
	rn = None # Right neighbor

	tl = None # Top-left neighbor
	tr = None # Top-right neighbor
	bl = None # Bottom-left neighbor
	br = None # Bottom-right neighbor

	T = None # Node Temperature (Symbolic variable)
	def __init__(self, idx) -> None:
		self.index = idx

# Initialize grid of nodes
columns = 8
rows = 3
grid = []
idx = 1
for j in range(rows):
	r = []
	for i in range(idx, idx + columns):
		node = TNode(i)
		r.append(node)
		idx += 1
	grid.append(r)

# Now that we have initialized a grid, we find all the neighbors
for i in range(columns):
	for j in range(rows):
		node  = grid[j][i]
		# Bottom & Top neighbors
		if node.index + columns <= rows * columns:
			# Find bottom and top neighbor pairs
			# Since these pairs are each others vertical neighbors,
			# we only need to scan once to locate both the top and bottom neighbors
			bottom_neighbor = grid[j+1][i]
			node.bn = bottom_neighbor
			bottom_neighbor.tn = node
		
		# Left & Right Neighbors
		if i + 1 <= 7:
			# Find right and left neighbor pairs
			# Since these pairs are each others horizontal neighbors,
			# we only need to scan once to locate both the right and left neighbors
			right_neighbor = grid[j][i+1]
			node.rn = right_neighbor
			right_neighbor.ln = node

		if j-1 >= 0 and i-1 >= 0:
			top_left = grid[j-1][i-1]
			node.tl = top_left
			top_left.br = node

		if j+1 <=2 and i-1 >= 0:
			bottom_left = grid[j+1][i-1]
			node.bl = bottom_left
			bottom_left.tr = node

# Create list of symbolic variables
# Axis of symmetry runs down the middle row
# Therefore, per column there are only 2 variables:
# 	"Inner" temperature and "Outer" temperature
# Then we only need to label these 2 variables per column

variable_set = []
for j in range(rows):
	for i in range(columns):
		node = grid[j][i]
		if j == 1:
			node.T = sympy.symbols("T_i_" + str(node.index))
		else:
			node.T = sympy.symbols("T_o_" + str(node.index))
		variable_set.append(node.T)

# Create the system of nodal finite-difference equations

equation_set = []
for i in range(columns):
	for j in range(rows):
		node  = grid[j][i]

		neighbors = [
			node.tn,
			node.bn,
			node.ln,
			node.rn,
			node.tr,
			node.tl,
			node.br,
			node.bl
		]
		num_neighbors = 0
		for value in neighbors:
			num_neighbors += int(value != None)

		if num_neighbors == 8:
			# This is an interior node
			eq = sympy.Eq(node.tn.T + node.bn.T + node.rn.T + node.ln.T - 4.0*node.T, 0.0)
			equation_set.append(eq)
		
		if num_neighbors == 5 and not node.index in [9, 16]:
			# This is a node at plane surface with convection
			# Find out if the convection surface is above or below
			if node.tn == None:
				# Convection surface is above
				SigmaTn = 2.0*node.bn.T + node.ln.T + node.rn.T
				# We precalculate a constant A to make our equations simpler
				A = 2.0*h*dx / k
				eq = sympy.Eq(SigmaTn + A*T_inf - A*node.T - 4.0*node.T, 0)
				equation_set.append(eq)
			else:
				# Convection surface is below
				SigmaTn = 2.0*node.tn.T + node.ln.T + node.rn.T
				# We precalculate a constant A to make our equations simpler
				A = 2.0*h*dx / k
				eq = sympy.Eq(SigmaTn + A*T_inf - A*node.T - 4.0*node.T, 0)
				equation_set.append(eq)
		else:
			# Handle edge cases
			if node.index == 16:
				eq = sympy.Eq(2.0*node.ln.T + node.tn.T + node.bn.T - 4.0*node.T, 0)
				equation_set.append(eq)
			elif node.index == 9:
				eq = sympy.Eq(node.tn.T + node.bn.T + node.rn.T + T_b - 4.0*node.T, 0)
				equation_set.append(eq)
		
		if num_neighbors == 3:
			# Adiabatic surface can be treated as axis of symmetry
			# This becomes similar to the nodes with 5 neighbors and convection
			# Determine if the convection surface is above or below
			
			# We precalculate a constant A to make our equations simpler
			A = 2.0*h*dx / k
			if node.tn == None:
				# Convection surface is above
				if node.index == 1:
					SigmaTn = 2.0*node.bn.T + T_b + node.rn.T
					eq = sympy.Eq(SigmaTn + A*T_inf - A*node.T - 4.0*node.T, 0)
					equation_set.append(eq)
				elif node.index == 8:
					SigmaTn = 2.0*node.bn.T + 2.0*node.ln.T
					eq = sympy.Eq(SigmaTn + A*T_inf - A*node.T - 4.0*node.T, 0)
					equation_set.append(eq)
			else:
				# Convection surface is below
				if node.index == 17:
					SigmaTn = 2.0*node.tn.T + T_b + node.rn.T
					eq = sympy.Eq(SigmaTn + A*T_inf - A*node.T - 4.0*node.T, 0)
					equation_set.append(eq)
				elif node.index == 24:
					SigmaTn = 2.0*node.tn.T + 2.0*node.ln.T
					eq = sympy.Eq(SigmaTn + A*T_inf - A*node.T - 4.0*node.T, 0)
					equation_set.append(eq)

# Solve the system and sort the node temperatures
# Then write them to an excel file

# with open("latex_equation_set.txt", "w") as f:
# 	interior_equations = []
# 	exterior_equations_conv = []
# 	exterior_equations_insul = []
# 	for eq in equation_set:
# 		line = "$$\n" + sympy.latex(eq) + "\n$$\n"
# 		indices = re.findall(r"T_{[io] (\d+)}", line)
# 		line = re.sub(r"T_{[io] (\d+)}", r"T_{\1}", line)
		
# 		if "h" in line:
# 			exterior_equations_conv.append(line)
# 		elif "8" in indices or "16" in indices or "24" in indices:
# 			exterior_equations_insul.append(line)
# 		else:
# 			interior_equations.append(line)
# 	f.write("\\subsection{Interior Equations}\n\n")
# 	f.writelines(interior_equations)
# 	f.write("\n\n\\subsection{Exterior Equations with Convection}\n\n")
# 	f.writelines(exterior_equations_conv)
# 	f.write("\n\n\\subsection{Exterior Equations with Insulation}\n\n")
# 	f.writelines(exterior_equations_insul)

solution = sympy.solve(equation_set, variable_set)

workbook = xlsxwriter.Workbook("NodeTemperatures.xlsx")
worksheet = workbook.add_worksheet()

for i in range(columns):
	for j in range(rows):
		node  = grid[j][i]
		worksheet.write(j,i,solution[node.T])

workbook.close()

df = pd.read_excel("NodeTemperatures.xlsx")
print(df)

input("\n\nPress Enter to close program")