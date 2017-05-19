#!C:\Python34
import sys
import CsvHandler
import os
import re
import csv
import WordHandler


class PathFinder:
	def __init__ (self):
		self.files = []

	def checkIfDirExists(self, directory):
		return os.path.isdir(directory)

	def checkIfFileExists(self, file):
		return os.path.isfile(file)

	def findFiles(self, directory):
		for root, dirs, files in os.walk(directory):
		    for file in files:
		        if file.endswith(".doc"):
		             # self.files.append(os.path.join(root, file))
		             self.files.append(file)

class PrettyPrint:
	def __init__ (self, consoleLength = 80):
		self.usedLength = 0
		self.consoleLength = consoleLength
		self.processText = ""

	def printLine(self, text):
		print ("[.] %s" % text)

	def printError(self, text):
		print ("[!] %s" % text)

	def printProcess(self, text):
		self.processText = text
		sys.stdout.write("[*] %s\r" % text)
		self.usedLength = len(text) 
		sys.stdout.flush()

	# [  OK  ]
	# [ FAIL ]

	def printOk(self):
		numOfWhiteSpace = self.consoleLength - self.usedLength - 8
		sys.stdout.write("[+] %s" % self.processText)
		sys.stdout.write(" " * numOfWhiteSpace)
		print("[  OK  ]")
		# sys.stdout.flush() 

	def printFail(self):
		numOfWhiteSpace = self.consoleLength - self.usedLength - 8
		sys.stdout.write("[-] %s" % self.processText)
		sys.stdout.write(" " * numOfWhiteSpace)
		print("[ FAIL ]")
		# sys.stdout.flush() 


if __name__ == "__main__":
	CSVFILE = "C:\\Users\\emrecan\\Documents\\PythonWin32Automation\\test.csv"

	cH = CsvHandler.CsvHandler()
	# PathFinder is the class that finds the .doc files in specified directories.
	pF = PathFinder()
	pp = PrettyPrint(50)

	pp.printProcess("Checking csv file")
	if not pF.checkIfFileExists(CSVFILE):
		pp.printFail()
		pp.printError("The csv file does not exist! Exitting!!!")
		sys.exit(10)
	pp.printOk()

	pp.printProcess("Reading csv file")
	cH.readCsvFile(CSVFILE)
	pp.printOk()

	for inputDir, outputDir in cH.dirList.items():
		pp.printLine("I will process " + inputDir)
		pp.printProcess("Checking if input dir exists")
		if not pF.checkIfDirExists(inputDir):
			pp.printFail()
			print ("%s is not present! Skipping..." % inputDir)
			continue
		
		pp.printOk()

		pp.printProcess("Checking if output dir exists")
		if not pF.checkIfDirExists(outputDir):
			pp.printFail()
			pp.printProcess("Creating output dir")
			try:
				os.makedirs(outputDir)
			except Exception, e:
				pp.printFail()
				pp.printLine("Cannot create output dir!")
				pp.printLine(str(e))
				pp.printLine("Skipping!")
				continue
		pp.printOk()

		pp.printProcess("Finding doc files")
		pF.findFiles(inputDir)
		numOfFiles = len(pF.files) 
		if numOfFiles == 0:
			pp.printFail()
			pp.printLine("No doc files found! Skipping this dir!")
			continue
		pp.printOk()

		currentFile = 1
		WH = WordHandler
		pp.printProcess("Starting MS Word Application")
		try:
			wordApp = WH.OpenWordApp()	# Starts the word application
		except:
			pp.printFail()
			pp.printLine("Error occured while starting MS Word App. Exitting!!!")
			sys.exit(10)
		pp.printOk()

		for file in pF.files:
			pp.printProcess("Processing file " + str(currentFile) + " of " + str(numOfFiles))
			currentFile += 1
			
			try:
				WH.OpenWordDocument(wordApp, inputDir + "\\" + file)	# opens the word document in the previously started word application.
			except:
				pp.printFail()
				pp.printLine("Error while opening " + inputDir + "\\" + file + ". Skipping!")
				continue
			
			table = wordApp.ActiveDocument.Tables(wordApp.ActiveDocument.Tables.Count) 	# Select the last table in the document
			# We want to put 'NT' to test steps with no
			# assigned requirements.
			# We need to search for a number or a letter
			# in the 'Requirements' column.
			# Since we only want to process empty cells, 
			# we will skip cells with values. 
			search=re.compile(r'[a-zA-Z0-9]').search									

			for row in range(3, table.Rows.Count+1):
				value = WH.GetCellValue(table, row, 5)
				if value is None:											# If cell has no value i.e. cell may be merged... 
					continue
				
				# if not search(value):										# Check if the cell is empty.
				if not search(WH.GetCellValue(table, row, 4)):			# If empty, check if the step has not run before.
					WH.SetCellValue(table, row, 4, "DIF")
			try:
				WH.SaveWordDocument(wordApp, outputDir + "\\" + file)
			except:
				pp.printFail()
				pp.printLine("Can't save file to " + outputDir + "\\" + file)
				pp.printLine("Skipping!")
				continue
			
			pp.printOk()
			WH.CloseWordDocument(wordApp)		

		WH.CloseWordApp(wordApp)				

	pp.printLine("Finished processing! Have a nice day...")

	sys.exit(0)