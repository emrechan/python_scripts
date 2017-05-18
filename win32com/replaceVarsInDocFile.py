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

    docFile = args.file.replace(".csv", ".doc")
    logging.debug(docFile)

    if not os.path.isfile(docFile):
        logging.critical(docfile + " not found! Exiting...")
        sys.exit(10)

    cp = CsvParser.CsvParser()
    cp.setFile(args.file)    
    cp.parse()
    rows = cp.getRows()

    msWord = MsWord.MsWord()
    try:
        msWord.startWordApp()
    except MsWordself.StartWordAppError as e:
        logging.critical(e)
        sys.exit(10)

    msWord.setFile(docFile)
    try:
        msWord.openDocFile()
    except MsWord.OpenDocFileError as e:
        logging.critical(e)
        logging.info("Skipping " + files[0])

    msWord.setTrackChangesOff()

    all_reqs = ""
    variables = {}
    requirements = {}

    logging.info("Parsing csv file.")
    for row in rows:
        option = row[0].upper()
        if option == "V":
            variables[row[1]] = row[2]
        elif option == "R":
            requirements[row[1]] = row[2]
        else:
            logging.error("Option " + option + " is not a valid option!")
            logging.error("Skipping!")

    logging.info("Replacing variables.")
    for var in variables:
        findWhat = "<"+var+">"
        logging.debug("Find What: " + findWhat)
        replaceWith = variables[var]
        logging.debug("Replace With: " + replaceWith)
        msWord.FindAndReplace(findWhat, replaceWith)

    logging.info("Replacing requirements.")
    for req in requirements:
        findWhat = "<?"+req+"?>"
        logging.debug("Find What: " + findWhat)
        replaceWith = requirements[req]
        logging.debug("Replace With: " + replaceWith)
        msWord.FindAndReplace(findWhat, replaceWith)
        all_reqs += requirements[req]

    findWhat = "<?all_reqs?>"
    logging.debug("Find What: " + findWhat)
    replaceWith = all_reqs
    logging.debug("Replace With: " + replaceWith)
    msWord.FindAndReplace(findWhat, replaceWith)

    msWord.setTrackChangesOn()
    
    if not msWord.save():
        logging.critical("Error saving file " + docFile)

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

    try:
        msWord.quitWordApp()
    except MsWord.QuitWordAppError as e:
        logging.critical(e)
        logging.critical("Unrecoverable error!!!")
        sys.exit(10)

    logging.info("All done!")

if __name__ == "__main__":
    main()
