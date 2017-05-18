#!c:\Python27\python2.7.exe

import modules.CsvParser as CsvParser
import modules.MsWord as MsWord
import argparse
import logging
import os.path
import sys

FORMAT = "%(asctime)s - %(levelname)s - %(message)s"
logging.basicConfig(format=FORMAT, level=logging.DEBUG)

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("-f", "--file" , required=True, help=".csv file to parse")
    args = parser.parse_args()

    if not os.path.isfile(args.file):
        logging.critical("File not found! Exiting...")
        sys.exit(10)

    
    cp = CsvParser.CsvParser()
    cp.setFile(args.file)    
    cp.parse()
    rows = cp.getRows()
    
    msWord = MsWord.MsWord()
    msWord.startWordApp()
    if not msWord.wordApp:
        logging.critical("Error opening Ms Word App! Exiting...")
        sys.exit(10)

    for row in rows:
        msWord.setFile(row[0])
        msWord.openDocFile()
        if not msWord.doc:
            logging.critical("Error opening file: " + row[0])
            logging.critical("Exiting...")
            msWord.quitWordApp()
            sys.exit(10)
        isFileSaved = msWord.saveFileAs(row[1])
        if not isFileSaved:
            logging.critical("Error saving file as: " + row[1])
            logging.critical("Skipping...")
        else:
            logging.info("File saved successfully as: " + row[1])
        isFileClosed = msWord.closeDocFile()
        if not isFileClosed:
            logging.critical("Error closing file " + row[0])
            logging.critical("Exiting...")
            msWord.quitWordApp()
            sys.exit(10)

    isWordAppClosed = msWord.quitWordApp()
    if not isWordAppClosed:
        logging.critical("Problem closing Word App!")
        exit(10)

    logging.info("All done!")

if __name__ == "__main__":
    main()