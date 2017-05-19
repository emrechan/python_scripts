#!c:\Python27\python2.7.exe

import modules.CsvParser as CsvParser
import modules.MsWord as MsWord
import argparse
import logging
import os.path
import sys

FORMAT = "%(asctime)s - %(levelname)s - %(message)s"
logging.basicConfig(format=FORMAT, level=logging.DEBUG)

def isFileExits(file):
    if not os.path.isfile(file):
        return False
    return True

def parseCsv(file):
    if not isFileExits(file):
        logging.critical("File not found! Exiting...")
        sys.exit(10)

    cp = CsvParser.CsvParser()
    cp.setFile(file)    
    cp.parse()
    rows = cp.getRows()
    return rows

def listFilesToCreateFromTemplate(rows):
    listOfFilesToCreate = []
    for row in rows:
        if row[2].upper() == "N":
            logging.debug("Skipping " + row[1])
            continue
        listOfFilesToCreate.append([row[0],row[1]])
    return listOfFilesToCreate

def checkIfConfigFilesExists(listOfFilesToCreate):
    # Config files checks
    filesWithExistingConfFile = []
    for files in listOfFilesToCreate:
        confFile = files[1] + ".csv"
        if not isFileExits(confFile):
            logging.warning("No configuration file found for " + files[1])
            continue
        filesWithExistingConfFile.append(files)
    return filesWithExistingConfFile

def doFileModificationOnTemplate(msWord,csvFile):
    msWord.setTrackChangesOff()

    all_reqs = ""
    variables = {}
    requirements = {}

    logging.info("Parsing csv file.")
    parsedCsvRows = parseCsv(csvFile)
    for row in parsedCsvRows:
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

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("-f", "--file" , required=True, help=".csv file to parse")
    args = parser.parse_args()

    parsedCsvRows = parseCsv(args.file)

    to_create = listFilesToCreateFromTemplate(parsedCsvRows)
    if not len(to_create) > 0:
        logging.info("No files to work with!")
        sys.exit(0)

    to_create = checkIfConfigFilesExists(to_create)

    if not len(to_create) > 0:
        logging.critical("No configration files available to work with!")
        logging.critical("Stoping...")
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
            confFile = files[1] + ".csv"
            doFileModificationOnTemplate(msWord, confFile)
            msWord.saveFileAs(files[1])
            logging.info("File saved successfully as: " + files[1])
            msWord.closeDocFile()
        except MsWord.OpenDocFileError as e:
            logging.critical(e)
            logging.info("Skipping " + files[0])
            continue
        except MsWord.SaveAsError as e:
            logging.critical(e)
            logging.critical("Couldn't save file " + files[1])            
        except MsWord.CloseDocFileError as e:
            logging.critical(e)
            logging.critical("Stoping rest of the process")
            try:
                msWord.quitWordApp()
            except MsWord.QuitWordAppError as e:
                logging.critical(e)
                logging.critical("Unrecoverable error!!!")
                sys.exit(10)
        except Exception as e:
            logging.critical("Unknown error: " + str(e))
            
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