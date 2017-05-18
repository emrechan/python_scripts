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

    to_create = []
    for row in rows:
        if row[2].upper() == "N":
            logging.debug("Skipping " + row[1])
            continue
        to_create.append([row[0],row[1]])
    
    if not len(to_create) > 0:
        logging.info("No files to work with!")
        sys.exit(0)


    msWord = MsWord.MsWord()    
    logging.debug("Starting MS Word App!")
    try:
        msWord.startWordApp()
    except MsWordself.StartWordAppError as e:
        logging.critical(e)
        sys.exit(10)

    for files in to_create:
        msWord.setFile(files[0])
        logging.debug("Opening " + files[0])
        try:
            msWord.openDocFile()
        except MsWord.OpenDocFileError as e:
            logging.critical(e)
            logging.info("Skipping " + files[0])
            continue
            
        try:
            msWord.saveFileAs(files[1])
            logging.info("File saved successfully as: " + files[1])
        except MsWord.SaveAsError as e:
            logging.critical(e)
            logging.critical("Couldn't save file " + files[1])            

        try: 
            msWord.closeDocFile()
        except MsWord.CloseDocFileError as e:
            logging.critical(e)
            logging.critical("Stoping rest of the process")
            try:
                msWord.quitWordApp()
            except MsWord.QuitWordAppError as e:
                logging.critical(e)
                logging.critical("Unrecoverable error!!!")
                sys.exit(10)

    logging.debug("Stopping MS Word App!")
    try:
        msWord.quitWordApp()
    except MsWord.QuitWordAppError as e:
        logging.critical(e)
        logging.critical("Unrecoverable error!!!")
        sys.exit(10)

    logging.info("All done!")

if __name__ == "__main__":
    main()