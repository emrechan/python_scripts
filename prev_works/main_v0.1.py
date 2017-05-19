#!C:\Python34
import win32com.client
import sys
import CsvHandler
import os
import re
import csv

def OpenWordApp():
	try:
		wordApp = win32com.client.DispatchEx("Word.Application")
		wordApp.Visible = False #Keep comment after tests
		wordApp.DisplayAlerts = False
		return wordApp
	except:
		print ("Error occured while accessing word application! Shutting down...")
		wordApp.Quit()
		sys.exit(1)

def OpenWordDocument(wordApp, fileLocation):
	try:
		wordApp.Documents.Open(fileLocation)
	except:
		print ("Error occured while openning document! Shutting down...")
		wordApp.ActiveDocument.Close()
		wordApp.Quit()
		sys.exit(1)

def FindAndReplace(wordApp, findWhat, replaceWith):
	wordApp.Selection.Find.Execute(findWhat, False, False, False, False, False, \
									True, 1, False, replaceWith, 2)

def SaveWordDocument(wordApp, fileLocation):
	try:
		wordApp.ActiveDocument.SaveAs(fileLocation)		
		print ("Successfully updated the document")
	except:
		print ("Error saving the document! Shutting down...")
		CloseWordDocument(wordApp)
		CloseWordApp(wordApp)
		sys.exit(2)

def GetCellValue(table, row, column):
	try:
		return table.Cell(Row=row, Column=column).Range.Text
	except:
		return None

def SetCellValue(table, row, column, text):
	try:
		table.Cell(Row=row, Column=column).Range.Text = text
	except:
		# print str(sys.exc_info()[0]) + ": " + str(sys.exc_info()[1])
		print ("Cannot write to Cell(%d,%d)" % (row, column))
		pass

def CloseWordDocument(wordApp):
	wordApp.ActiveDocument.Close()

def CloseWordApp(wordApp):
	wordApp.Quit()

if __name__ == "__main__":
	
	# csv = CsvHandler
	# csv.readCsvFile("test.csv")
	wordApp = OpenWordApp()
	OpenWordDocument(wordApp, "C:\\Users\\emrecan\\Documents\\PythonWin32Automation\\T_VOL7_ES_SYST_LINK11B_001-001_SSSB-D IZMIR.doc")
	currentDocument = wordApp.ActiveDocument
	table = currentDocument.Tables(currentDocument.Tables.Count) # Select the last table in the document
	# print table.Columns.Count
	# print table.Rows.Count
	search=re.compile(r'[a-zA-Z0-9]').search
	for row in range(3, table.Rows.Count+1):
		value = GetCellValue(table, row, 5)
		if value is None:
			continue
		
		if not search(value):
			if not search(GetCellValue(table, row, 4)):
				print (GetCellValue(table, row, 1).rstrip())
				print (GetCellValue(table, row, 4).rstrip())
				SetCellValue(table, row, 4, "NT")
	
	# FindAndReplace(wordApp, "<var1>", "val1")
	SaveWordDocument(wordApp, "C:\\Users\\emrecan\\Documents\\PythonWin32Automation\\T_VOL7_ES_SYST_LINK11B_001-001_SSSB-D IZMIR_modified.doc")
	CloseWordDocument(wordApp)
	CloseWordApp(wordApp)
	sys.exit(0)