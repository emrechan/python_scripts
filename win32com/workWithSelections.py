#!c:\Python27\python2.7.exe
import modules.MsWord as MsWord
import argparse
import logging
import os.path
import sys
import re

FORMAT = "%(asctime)s - %(levelname)s - %(message)s"
logging.basicConfig(format=FORMAT, level=logging.INFO)

def findTest(curTable):
    try:
        curText = curTable.Cell(6,6).Range.Text
        for match in re.finditer(r'<if(.*?)fi>', curText):
            if "<if false" in match.group(0):
                curText = curText.replace(match.group(0), "")
            else:
                curText = curText.replace(match.group(0), match.group(1))
            curTable.Cell(6,6).Range.Text = curText
    except Exception as e:
        print str(e)
        pass

def workWithSelections(docFile):
    try:
        msWord = MsWord.MsWord()
        logging.debug("Starting MS Word App!")
        msWord.startWordApp()
        msWord.setFile(docFile)
        logging.debug("Opening " + docFile)
        msWord.openDocFile()
        msWord.setTrackChangesOff()
        findTest(msWord.getTable(2))
        msWord.setTrackChangesOn()
        msWord.save()
        msWord.closeDocFile()
    except MsWord.StartWordAppError as e:
        logging.critical(e)
        sys.exit(10)
    except MsWord.OpenDocFileError as e:
        logging.critical(e)
        logging.info("Cannot open  " + docFile)
    except MsWord.TableIndexNotFoundError as e:
        logging.critical(e)
        logging.critical("Cannot find the specified table")
    except MsWord.CloseDocFileError as e:
        logging.critical(e)
        logging.critical("Cannot close the doc file")
    except Exception as e:
        logging.critical(e)
        logging.critical("Unknown error: " + str(e))
    finally:
        logging.debug("Stopping MS Word App!")
        try:
            msWord.quitWordApp()
        except MsWord.QuitWordAppError as e:
            logging.critical(e)
            logging.critical("Unrecoverable error!!!")
            sys.exit(10)

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("-f", "--file" , required=True, help=".doc file to parse")
    args = parser.parse_args()

    workWithSelections(args.file)

if __name__ == "__main__":
    main()
