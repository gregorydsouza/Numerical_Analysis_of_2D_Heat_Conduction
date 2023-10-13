import sympy

# Properties
k = 8 # Btu / hr*ft*degF
h = 96 # Btu / hr*ft*ft*degF
T_inf = 80 # degF
T_b = 200 # degF

# Create a class to generalize each node
class TNode:
	# Neighbors
	bn = None
	tn = None
	ln = None
	rn = None
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

# Create list of symbolic variables
# Axis of symmetry runs down the middle row
# Therefore, per column there are only 2 variables:
# 	"Inner" temperature and "Outer" temperature
# Then we only need to label these 2 variables per column

sym_string = ""
for j in range(rows - 1):
	for i in range(columns):
		if j == 1:
			sym_string += "T_i_" + str(i + 1) + ", "
		else:
			sym_string += "T_o_" + str(i + 1) + ", "
sym_string = sym_string.rstrip(", ")

sympy.symbols(sym_string)
del(sym_string)

# Create the system of nodal finite-difference equations