import os.path
import logging
import win32com.client as win32


FORMAT = "%(asctime)s - %(levelname)s - %(message)s"
logging.basicConfig(format=FORMAT, level=logging.DEBUG)

class Error(Exception):
    """Base class for other exceptions"""
    pass

class StartWordAppError(Error):
    """Raised when the input value is too small"""
    def __init__(self, e):
        Error.__init__(self,e[2])
    pass

class OpenDocFileError(Error):
    """Raised when the input value is too small"""
    def __init__(self, e):
        Error.__init__(self,e[2])

class SaveAsError(Error):
    def __init__(self, e):
        Error.__init__(self,e[2])

class CloseDocFileError(Error):
    def __init__(self, e):
        Error.__init__(self,e[2])

class QuitWordAppError(Error):
    def __init__(self, e):
        Error.__init__(self,e[2])

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
        except Exception as e:
            raise StartWordAppError(e)

    def setFile(self, file):
        self.file = file

    def openDocFile(self):
        try:
            self.doc = self.wordApp.Documents.Open(self.file)
        except Exception as e:
            raise OpenDocFileError(e)

    def findTestStepsTable(self):
        if not self.doc.Tables.Count == 2:
            return False
        else:
            self.table = self.doc.Tables.Item(2)
            return True

    def findTableRowsCount(self):
        if self.table:
            return self.table.Rows.Count

    def Find(self, findWhat):
        return self.wordApp.Selection.Find.Execute(findWhat, False, False, False, False, \
                                                   False, True, 1, False, False, 0)
    # def findAllRows(self, findWhat):
        
    def Selection(self):
        return self.wordApp.Selection

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
        except Exception as e:
            raise SaveAsError(e)

    def save(self):
        try:
            self.doc.Save()
            return True
        except:
            return False

    def closeDocFile(self):
        try:
            self.doc.Close(SaveChanges=False)
        except Exception as e:
            raise CloseDocFileError(e)

    def quitWordApp(self):
        try:
            self.wordApp.Application.Quit()
        except Exception as e:
            raise QuitWordAppError(e)
