#!/usr/bin/env python3

##################################################
## testy.py
##################################################
## Author: clenahan@cloud2gnd.com
## Copyright: Copyright 2024
## Version: 1.0 - Aug 2024 - Initial Release
##################################################


import getopt, sys
import tcrl
import iopimport
import tca
import extract
import tcmt
from mscgen import processMyDocx

def showHelp():
    install = "install instructions:\npython3 -m pip install python-docx\npython3 -m pip install openpyxl\npython3 -m pip install nltk\npython3 -m nltk.downloader punkt\npython3 -m nltk.downloader punkt_tab\n\n"
    print (install + "uasge:  python3 -m testy [-h]\n\tpython3 -m testy [-v]\n\tpython3 -m testy [-t test_file [-n]]  \t\tTC extraction\n\tpython3 -m testy [-i iop_test_plan]\t\tIOP Test Tool import generator\n\tpython3 -m testy [-s spec]\t\t\tTCA generator\n\tpython3 -m testy [-m my_docx]\t\t\t teststeps2msc\n\tpython3 -m testy [-x my_docx]\t\t\t Extract Test Steps\t\t\t teststeps2msc\n\tpython3 -m testy [-c my_docx]\t\t\t Extract TCMT")

kTCRL = 1
kIOP = 2
kTCA = 3
kMYFUNC = 4  # New mode for -m
kEXTRACT = 5
kTCMT = 6

if __name__ == "__main__":
    argumentList = sys.argv[1:]
    options ="ht:i:s:nm:x:c:"
    path=""
    mode = None
    naturalSort = False
    try:
        arguments, values = getopt.getopt(argumentList,options)
        for currentArgument, currentValue in arguments:
            if currentArgument in ("-h"):
                showHelp()
            elif currentArgument in ("-t"):
                path = currentValue
                mode = kTCRL
            elif currentArgument in ("-i"):
                path = currentValue
                mode = kIOP
            elif currentArgument in ("-s"):
                path = currentValue
                mode = kTCA
            elif currentArgument in ("-m"):
                path = currentValue
                mode = kMYFUNC
            elif currentArgument in ("-x"):
                print("-x")
                path = currentValue
                mode = kEXTRACT
            elif currentArgument in ("-c"):
                print("-c")
                path = currentValue
                mode = kTCMT
            elif currentArgument in ("-n"):
                naturalSort =True
    except getopt.error as err:
        print (str(err))

    if  mode == kTCRL:
        tcrl.dumpTcrl(path,naturalSort)
    elif  mode == kIOP:
        iopimport.processIOPTestPlan(path)
    elif  mode == kTCA:
        tca.processTCA(path)
    elif  mode == kMYFUNC:
        processMyDocx(path)
    elif  mode == kEXTRACT:
        extract.extractTS(path)
    elif  mode == kTCMT:
        tcmt.extractTCMT(path)
    else:
        showHelp()
