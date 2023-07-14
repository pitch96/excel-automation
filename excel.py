from openpyxl import Workbook
from openpyxl.utils import FORMULAE

def writeExcel(data):
	# Load the workbook
	workbook = Workbook()

	# Get the worksheet
	worksheet = workbook.active

	for row in data:
		worksheet.append(row)

	workbook.save('output.xlsx')