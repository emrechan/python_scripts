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
    logging.debug("Starting MS Word App!")
    if not msWord.startWordApp():
        logging.critical("Error opening Ms Word App! Exiting...")
        sys.exit(10)

    for row in rows:
        if row[2].upper() == "N":
            logging.debug("Skipping " + row[1])
            continue
        msWord.setFile(row[0])
        logging.debug("Opening " + row[0])
        msWord.openDocFile()
        if not msWord.openDocFile():
            logging.critical("Error opening file: " + row[0])
            logging.critical("Exiting...")
            msWord.quitWordApp()
            sys.exit(10)
        if not msWord.saveFileAs(row[1]):
            logging.critical("Error saving file as: " + row[1])
            logging.critical("Skipping...")
        else:
            logging.info("File saved successfully as: " + row[1])
        if not msWord.closeDocFile():
            logging.critical("Error closing file " + row[0])
            logging.critical("Exiting...")
            msWord.quitWordApp()
            sys.exit(10)
        logging.debug("Finished " + row[1])

    logging.debug("Stopping MS Word App!")
    if not msWord.quitWordApp():
        logging.critical("Problem closing Word App!")
        exit(10)

    logging.info("All done!")

if __name__ == "__main__":
    main()