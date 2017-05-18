import os.path
import logging
import win32com.client as win32


FORMAT = "%(asctime)s - %(levelname)s - %(message)s"
logging.basicConfig(format=FORMAT, level=logging.DEBUG)

class MsWord():
    def __init__(self):        
        self.wordApp = None
        self.doc = None
        self.file = None
        self.table = None

    def startWordApp(self):
        try:
            self.wordApp = win32.gencache.EnsureDispatch('Word.Application')
        except:
            pass

    def setFile(self, file):
        self.file = file

    def openDocFile(self):
        try:
            self.doc = self.wordApp.Documents.Open(self.file)
        except:
            pass

    def findTestStepsTable(self):
        if not self.doc.Tables.Count == 2:
            return False
        else:
            self.table = self.doc.Tables.Item(2)
            return True

    def findTableRowsCount(self):
        if self.table:
            return self.table.Rows.Count

    def saveFileAs(self, file):
        try:
            self.doc.SaveAs(file)
            return True
        except:
            return False

    def closeDocFile(self):
        try:
            self.doc.Close(SaveChanges=False)
            return True
        except:
            return False

    def quitWordApp(self):
        try:
            self.wordApp.Application.Quit()
            return True
        except:
            return False
