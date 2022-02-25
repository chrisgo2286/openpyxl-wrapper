import openpyxl

def load_excel_file(file_name, worksheet=None, data_only=False):
	""" Returns workbook and worksheet """
	if data_only == False:
		wb = openpyxl.load_workbook(file_name)
	else:
		wb = openpyxl.load_workbook(file_name, data_only=True)
	ws = wb[worksheet] if worksheet else wb.active
	return wb, ws

def save_and_close(workbook, file_name):
	""" Saves the file and closes workbooks """
	workbook.save(file_name)
	workbook.close()
	