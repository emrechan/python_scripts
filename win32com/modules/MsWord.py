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
            self.wordApp.Visible = False #Keep comment after tests
            self.wordApp.DisplayAlerts = False
            return True
        except:
            return False

    def setFile(self, file):
        self.file = file

    def openDocFile(self):
        try:
            self.doc = self.wordApp.Documents.Open(self.file)
            return True
        except Exception as e:
            logging.critical(e)
            return False

    def findTestStepsTable(self):
        if not self.doc.Tables.Count == 2:
            return False
        else:
            self.table = self.doc.Tables.Item(2)
            return True

    def findTableRowsCount(self):
        if self.table:
            return self.table.Rows.Count

    def FindAndReplace(self, findWhat, replaceWith):
        self.wordApp.Selection.Find.Execute(findWhat, False, False, False, False, False, \
                                            True, 1, False, replaceWith, 2)

    def setTrackChangesOff(self):
        self.doc.TrackRevisions = False

    def setTrackChangesOn(self):
        self.doc.TrackRevisions = True            

    def saveFileAs(self, file):
        try:
            self.doc.SaveAs(file)
            return True
        except:
            return False

    def save(self):
        try:
            self.doc.Save()
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
