#!C:\Python34
import sys
import os
import re
import csv
import WordHandler
import unicodedata

class CsvHandler:
	def __init__ (self):
		self.dirList = {}
		
	def readCsvFile(self, fileLocation):
		with open(fileLocation, 'rb') as csvfile:
			reader = csv.DictReader(csvfile, delimiter=';', quotechar='|')		
			
			for row in reader:
				if not row['InputFolder'] is "" and not row['TcType'] is "":
					self.dirList[row['InputFolder']] = row['TcType']

class PathFinder:
	def __init__ (self):
		self.files = []

	def checkIfDirExists(self, directory):
		return os.path.isdir(directory)

	def checkIfFileExists(self, file):
		return os.path.isfile(file)

	def findFiles(self, directory):
		self.files = []		# clear the list if it contains previous data  
		for root, dirs, files in os.walk(directory):
		    for file in files:
		    	if "~" in file:
		    		continue
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
	CSVFILE = "C:\\Users\\emrecan\\Documents\\PythonWin32Automation\\test2.csv"
	stepStatusDict = {
	'OK':0,
	'NOK':0,
	'NT':0,
	'DIF':0,
	'OTHER':0,
	'VWC':0
	}

	cH = CsvHandler()
	# PathFinder is the class that finds the .doc files in specified directories.
	pF = PathFinder()
	pp = PrettyPrint(50)

	fh = open('output.csv', 'w')
	fhSt = open('statistics.csv', 'w')

	fhSt.write("%s;%s;%s;%s;%s;%s;%s;%s\n" %("TC Type","File Name","OK","NOK","NT","DIF","VWC","OTHER"))
	fh.write("%s; %s;%s\n" % ("File Name","Step No","Step Status"))

	pp.printProcess("Checking csv file")
	if not pF.checkIfFileExists(CSVFILE):
		pp.printFail()
		pp.printError("The csv file does not exist! Exitting!!!")
		sys.exit(10)
	pp.printOk()

	pp.printProcess("Reading csv file")
	cH.readCsvFile(CSVFILE)
	pp.printOk()

	WH = WordHandler
	pp.printProcess("Starting MS Word Application")
	try:
		wordApp = WH.OpenWordApp()	# Starts the word application
	except:
		pp.printFail()
		pp.printLine("Error occured while starting MS Word App. Exitting!!!")
		sys.exit(10)
	pp.printOk()


	for inputDir, tcType in cH.dirList.items():
		pp.printLine("I will process " + inputDir)
		pp.printProcess("Checking if input dir exists")
		if not pF.checkIfDirExists(inputDir):
			pp.printFail()
			print ("%s is not present! Skipping..." % inputDir)
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

		for file in pF.files:
			pp.printProcess("Processing file " + str(currentFile) + " of " + str(numOfFiles))
			currentFile += 1
			
			try:
				WH.OpenWordDocument(wordApp, inputDir + "\\" + file)	# opens the word document in the previously started word application.
			except:
				pp.printFail()
				pp.printLine("Error while opening " + inputDir + "\\" + file + ". Skipping!")
				continue
			
			table = wordApp.ActiveDocument.Tables(2) 	# Select the appropriate table.
			# We want to put 'NT' to test steps with no
			# assigned requirements.
			# We need to search for a number or a letter
			# in the 'Requirements' column.
			# Since we only want to process empty cells, 
			# we will skip cells with values. 
			for row in range(3, table.Rows.Count+1):
				value = WH.GetCellValue(table, row, 4)
				if value is None:											# If cell has no value i.e. cell may be merged... 
					continue
				stepNo = WH.GetCellValue(table, row, 1)
				stepStatus = WH.GetCellValue(table, row, 4)
				stepNoCleaned = re.sub('[\000-\040]', '', stepNo)			# stepNo value includes \n and BEL chars. We remove them.
				stepStatusClean = re.sub('[\000-\040]', '', stepStatus)		# stepStatus value includes \n and BEL chars. We remove them.
				if stepNoCleaned == "" and stepStatusClean == "":			# Sometimes both the stepNo and stepStatus values are empty strings (maybe merged rows?). We ignore them. 
					continue
				if stepStatusClean in stepStatusDict:
					stepStatusDict[stepStatusClean] += 1
				elif stepStatusClean == "":
					stepStatusDict['NT'] += 1
				else:
					stepStatusDict['OTHER'] += 1
				fh.write("%s; %s;%s\n" % (file,stepNoCleaned,stepStatusClean))
			
			pp.printOk()
			WH.CloseWordDocument(wordApp)
			fhSt.write("%s;%s;%d;%d;%d;%d;%d;%d\n" %(tcType,file,stepStatusDict['OK'],stepStatusDict['NOK'],
				stepStatusDict['NT'],stepStatusDict['DIF'],stepStatusDict['VWC'],stepStatusDict['OTHER']))
			for key in stepStatusDict:
				stepStatusDict[key] = 0

	WH.CloseWordApp(wordApp)				
	pp.printLine("Finished processing! Have a nice day...")

	fh.close()
	fhSt.close()
	sys.exit(0)