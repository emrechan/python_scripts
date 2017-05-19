#!C:\Python34
import csv

class CsvHandler:
	def __init__ (self):
		self.dirList = {}
		
	def readCsvFile(self, fileLocation):
		with open(fileLocation, 'rb') as csvfile:
			reader = csv.DictReader(csvfile, delimiter=';', quotechar='|')		
			
			for row in reader:
				if not row['InputFolder'] is "" and not row['OutputFolder'] is "":
					self.dirList[row['InputFolder']] = row['OutputFolder']