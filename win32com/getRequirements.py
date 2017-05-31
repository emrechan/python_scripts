#!c:\Python27\python2.7.exe
import modules.MsWord as MsWord
import argparse
import logging
import os.path
import sys
import re

FORMAT = "%(asctime)s - %(levelname)s - %(message)s"
logging.basicConfig(format=FORMAT, level=logging.INFO)
paraMark = u'\r\r\x07'


def findRequirements(table):
    maxRows = table.Rows.Count + 1
    allReqs = {}
    for row in range(3,maxRows):
        try:
            text = table.Cell(row,5).Range.Text
            # for match in re.finditer(r'(\w+\d+/?\d?-?\w?)+',text, re.DOTALL):
            for match in re.finditer(r'(\w+\d+-?\w{0,2}-?\w{0,2}/?\d?-?)+',text, re.DOTALL):
                allReqs[match.group(0)] = 1
            # if re.search(r'[a-zA-Z0-9]', text, re.MULTILINE):
            #     text = linuxTr(text)
            #     allReqs[text] = 1
        except Exception as e:
            # logging.debug(e)
            pass
    return allReqs

def findLocation(doc):
    para = doc.Paragraphs.First
    tryNo = 1
    while tryNo < 10:
        # L16 location
        # match = re.match(r'.*located at (\w+).*', para.Range.Text)
        # AWCIES Sensors location
        match = re.match(r'.*Verification of the ARS Eskisehir Interface Exchange Capability with (.*)\.', para.Range.Text)
        if match:
            return match.group(1)
        para = para.Next()
        tryNo += 1

    return "not found"

def getRequirements(docFile):
    try:
        msWord = MsWord.MsWord()
        logging.debug("Starting MS Word App!")
        msWord.startWordApp()
        msWord.setFile(docFile)
        logging.debug("Opening " + docFile)
        msWord.openDocFile()
        # TODO Add the logic here
        allReqs = findRequirements(msWord.getTable(2))
        location = findLocation(msWord.getDocument())
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

        fileName = os.path.basename(docFile)
        # print location
        try:
            for req in allReqs:
                print fileName + ";" + location + ";" + req
        except UnboundLocalError:
            logging.critical("Couldn't find requirements in " + fileName)
def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("-f", "--file" , required=True, help=".doc file to parse")
    args = parser.parse_args()

    getRequirements(args.file)

if __name__ == "__main__":
    main()
