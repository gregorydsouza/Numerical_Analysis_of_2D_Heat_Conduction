import sympy
import xlsxwriter

# Properties
k = 8.0 # Btu / hr*ft*degF
h = 96.0 # Btu / hr*ft*ft*degF
T_inf = 80.0 # degF
T_b = 200.0 # degF
dx = 1.0 / 8.0 # inch
dy = dx # inch

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
				eq = sympy.Eq((2.0*node.bn.T + node.ln.T + node.rn.T) + (2.0*h*dx*T_inf / k), (2.0*node.T)*((h*dx / k) + 2.0))
				equation_set.append(eq)
			else:
				# Convection surface is below
				eq = sympy.Eq((2.0*node.tn.T + node.ln.T + node.rn.T) + (2.0*h*dx*T_inf / k), (2.0*node.T)*((h*dx / k) + 2.0))
				equation_set.append(eq)
		else:
			# Handle edge cases
			if node.index == 16:
				eq = sympy.Eq((2.0*node.ln.T + node.tn.T + node.bn.T) - 4.0*node.T, 0)
				equation_set.append(eq)
			elif node.index == 9:
				eq = sympy.Eq(node.tn.T + node.bn.T + node.rn.T + T_b - 4.0*node.T, 0)
				equation_set.append(eq)
		
		if num_neighbors == 3:
			# Adiabatic surface can be treated as axis of symmetry
			# This becomes similar to the nodes with 5 neighbors and convection
			# Determine if the convection surface is above or below
			if node.tn == None:
				# Convection surface is above
				if node.index in [1,17]:
					eq = sympy.Eq((2.0*node.bn.T + T_b + node.rn.T) + (2.0*h*dx*T_inf / k), (2.0*node.T)*((h*dx / k) + 2.0))
					equation_set.append(eq)
				else:
					eq = sympy.Eq((2.0*node.bn.T + 2.0*node.ln.T) + (2.0*h*dx*T_inf / k), (2.0*node.T)*((h*dx / k) + 2.0))
					equation_set.append(eq)
			else:
				# Convection surface is below
				if node.index in [1,17]:
					eq = sympy.Eq((2.0*node.tn.T + T_b + node.rn.T) + (2.0*h*dx*T_inf / k), (2.0*node.T)*((h*dx / k) + 2.0))
					equation_set.append(eq)
				else:
					eq = sympy.Eq((2.0*node.tn.T + 2.0*node.ln.T) + (2.0*h*dx*T_inf / k), (2.0*node.T)*((h*dx / k) + 2.0))
					equation_set.append(eq)

# Solve the system and sort the node temperatures
# Then write them to an excel file
solution = sympy.solve(equation_set, variable_set)

workbook = xlsxwriter.Workbook("NodeTemperatures.xlsx")
worksheet = workbook.add_worksheet()

for i in range(columns):
	for j in range(rows):
		node  = grid[j][i]
		worksheet.write(j,i,solution[node.T])

workbook.close()