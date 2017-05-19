from win32com.client import Dispatch
import os,sys


xlBottom = -4107
xlCenter = -4108
xlLeft = -4131

class ExcelWriter(object):
	"""Excel class for creating spreadsheets - esp writing data and formatting them
	Based in part on #http://snippets.dzone.com/posts/show/2036,
	and http://www.markcarter.me.uk/computing/python/excel.html
	"""
	def __init__(self, file_name, make_visible=False):
		"""Open spreadsheet"""
		self.excelapp = Dispatch("Excel.Application")
		if make_visible:
			self.excelapp.Visible = 1 #fun to watch!
		try:
			self.excelapp.Workbooks.Open(file_name)
		except:
			sys.exit(-1)
		self.workbook = self.excelapp.ActiveWorkbook
		self.file_name = file_name
		self.default_sheet = self.excelapp.ActiveSheet
	def list_sheet_names(self):		
		for sheet in self.workbook.Sheets:
			print (sheet.Name)
	def get_cell_value(self, row=1, column=1, sheet=None):
		if sheet == None:
			sheet = self.default_sheet
		return sheet.Cells(row,column).Value
	def set_cell_value(self, row=1, column=1, content="", sheet=None):
		if sheet == None:
			sheet = self.default_sheet
		sheet.Cells(row,column).Value = content
	def get_sheet(self, sheet_name):
		"""
		Get sheet by name.
		"""
		return self.workbook.Sheets(sheet_name)
	def activate_sheet(self, sheet_name):
		"""
		Activate named sheet.
		"""
		sheets = self.workbook.Sheets
		sheets(sheet_name).Activate() #http://mail.python.org/pipermail/python-win32/2002-February/000249.html
		self.default_sheet = self.excelapp.ActiveSheet
	def format_cell(self, row, column, style):
		sheet = self.default_sheet
		if style == "DEFAULT":
			sheet.Cells(row,column).Font.Bold = False
			sheet.Cells(row,column).Font.Name = "Arial"
			sheet.Cells(row,column).Font.Size = 11
			sheet.Cells(row,column).VerticalAlignment = xlBottom
			sheet.Cells(row,column).HorizontalAlignment = xlBottom
	def save(self):
		"""Save spreadsheet as filename - wipes if existing"""
		if os.path.exists(self.file_name):
			os.remove(self.file_name)
		self.workbook.SaveAs(self.file_name)
	def save_as(self, file_name):
		"""Save spreadsheet as filename - wipes if existing"""
		if os.path.exists(file_name):
			os.remove(file_name)
		self.workbook.SaveAs(file_name)
	def close(self):
			"""Close spreadsheet resources"""
			self.workbook.Saved = 0 #p.248 Using VBA 5
			self.workbook.Close(SaveChanges=0) #to avoid prompt
			self.excelapp.Quit()
			self.excelapp.Visible = 0 #must make Visible=0 before del self.excelapp or EXCEL.EXE remains in memory.
			del self.excelapp

if __name__ == "__main__":
	print ("opened")
	ew = ExcelWriter("D:\\tmp\\TC_KO_son.xlsx")
	ew.activate_sheet("ATMGW-SAT-GL-002")	

	title_counter = 0
	row_counter = 1
	isFirst = True

	for i in range(35,500):
		if ew.get_cell_value(i, 3) is None:
			if ew.get_cell_value(i, 2) is not None:
				# print (str(title_counter) + ". " + 
				# 	   ew.get_cell_value(i, 2)[
				# 	   							ew
				# 	   							.get_cell_value(i, 2)
				# 	   							.index(" ") + 1 : 
				#    								]
				# 	  )
				value = ew.get_cell_value(i, 2)
				index = value.index(" ")
				if not isFirst:
					title_counter += 1
				value = str(title_counter) + ". " + value[index+1:]
				value = value[index+1:]
				ew.set_cell_value(i, 2, value)
				 
				isFirst = False
				row_counter = 1
		else:
			value = str(title_counter) + "-" + str(row_counter)
			ew.set_cell_value(i, 2, value)
			ew.format_cell(i, 2, "DEFAULT")

	# print(ew.get_cell_value(35,2))
	# print(ew.get_cell_value(35,3))
	ew.save_as("D:\\tmp\\test.xlsx")
	ew.close()
	print ("closed")