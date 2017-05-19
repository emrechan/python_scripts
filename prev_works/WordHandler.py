#!C:\Python34
import win32com.client

def OpenWordApp():
	try:
		wordApp = win32com.client.DispatchEx("Word.Application")
		wordApp.Visible = False #Keep comment after tests
		wordApp.DisplayAlerts = False
		return wordApp
	except:
		# print ("Error occured while accessing word application! Shutting down...")
		wordApp.Quit()
		raise

def OpenWordDocument(wordApp, fileLocation):
	try:
		wordApp.Documents.Open(fileLocation)
	except:
		# print ("Error occured while openning document!")
		wordApp.ActiveDocument.Close()
		raise

def FindAndReplace(wordApp, findWhat, replaceWith):
	wordApp.Selection.Find.Execute(findWhat, False, False, False, False, False, \
									True, 1, False, replaceWith, 2)

def TurnOnTrackChanges(wordApp):
	wordApp.ActiveDocument.TrackRevisions = True

def SaveWordDocument(wordApp, fileLocation):
	try:
		wordApp.ActiveDocument.SaveAs(fileLocation)		
		# print ("Successfully updated the document")
	except:
		# print ("Error saving the document!")
		CloseWordDocument(wordApp)
		raise

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
		# print ("Cannot write to Cell(%d,%d)" % (row, column))
		pass

def CloseWordDocument(wordApp):
	wordApp.ActiveDocument.Close(False)

def CloseWordApp(wordApp):
	wordApp.Quit()