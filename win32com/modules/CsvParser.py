import csv

class CsvParser():
    def __init__(self):
        self.file = None
        self.spamreader = None
        self.rows = []
        self.header = True

    def setFile(self, file):
        self.file = file

    def hasHeader(self, header):
        self.header = header

    def parse(self):
        if not self.file:
            return None

        with open(self.file, 'rb') as csvfile:
            self.spamreader = csv.reader(csvfile, delimiter=';', quotechar='"')
            for row in self.spamreader:
                if self.header:             # Skip header line
                    self.header = False
                    continue
                self.rows.append(row)

    def getRows(self):
        return self.rows

