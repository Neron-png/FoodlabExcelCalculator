from openpyxl import Workbook


def saturated(ws, new_ws, work_area, max_row):

	sum = "=("
	for i in range(work_area[1], work_area[2] + 1):
		if new_ws['A' + str(i)].value is not None:
			if ':0' in new_ws['A' + str(i)].value:
				sum += "C" + str(i) + "+"

	if sum != "=(":
		sum = sum[:-1:] + ")"
	else:
		sum = 'empty'

	# for i in range(work_area[2] + 1, max_row + 1):
	# 	if ws[work_area[0] + str(i)].value == "saturated":
	# 		new_ws[chr(ord(work_area[0]) + 1) + str(i)] = sum
	# 		break
	# else:
	EOF = max_row + 2
	new_ws[work_area[0] + str(EOF)] = "Saturated"
	new_ws[chr(ord(work_area[0]) + 1) + str(EOF)] = sum


def mono(ws, new_ws, work_area, max_row):
	sum = "=("
	for i in range(work_area[1], work_area[2] + 1):
		if new_ws['A' + str(i)].value is not None:
			if ':1' in new_ws['A' + str(i)].value:
				sum += "C" + str(i) + "+"

	if sum != "=(":
		sum = sum[:-1:] + ")"
	else:
		sum = 'empty'

	EOF = max_row + 4
	new_ws[work_area[0] + str(EOF)] = "Mono"
	new_ws[chr(ord(work_area[0]) + 1) + str(EOF)] = sum

def poly(ws, new_ws, work_area, max_row):
	sum = "=("
	for i in range(work_area[1], work_area[2] + 1):

		cell = new_ws['A' + str(i)].value

		if cell is not None:
			if ":2" in cell or ":3" in cell or ":4" in cell:
				sum += "C" + str(i) + "+"

	if sum != "=(":
		sum = sum[:-1:] + ")"
	else:
		sum = 'empty'

	EOF = max_row + 6
	new_ws[work_area[0] + str(EOF)] = "Poly"
	new_ws[chr(ord(work_area[0]) + 1) + str(EOF)] = sum

def trans(ws, new_ws, work_area, max_row):
	sum = "=("
	for i in range(work_area[1], work_area[2] + 1):

		cell = new_ws['A' + str(i)].value

		if cell is not None:
			if "trans" in cell:
				sum += "C" + str(i) + "+"

	if sum != "=(":
		sum = sum[:-1:] + ")"
	else:
		sum = 'empty'

	EOF = max_row + 8
	new_ws[work_area[0] + str(EOF)] = "Trans"
	new_ws[chr(ord(work_area[0]) + 1) + str(EOF)] = sum