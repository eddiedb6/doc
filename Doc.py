# -*- coding:utf-8 -*-

import os
import sys
import shutil

from docx import Document
from docx.shared import Inches

sys.path.append(os.path.join(os.path.split(os.path.realpath(__file__))[0], "."))

from DocProcessor import *

if len(sys.argv) < 3:
    print("Usage: ")
    print("    python Doc.py DOC_PATH REG_1, ..., REG_n")
    sys.exit()

docPath = sys.argv[1]
if not os.path.exists(docPath):
    print("No file: " + docPath)
    sys.exit()

docPathArray = os.path.splitext(docPath)
if len(docPathArray) != 2:
    print("Invalid doc file path format: " + docPath)
    sys.exit()
nameIndexStr = ""
index = 0
docBakPath = ""
while True:
    docBakPath = docPathArray[0] + ".bak" + nameIndexStr + docPathArray[1]
    if not os.path.exists(docBakPath):
        break
    index += 1
    nameIndexStr = str(index)
shutil.copy2(docPath, docBakPath)    

for i in range(len(sys.argv) - 2):
    inputArg = sys.argv[i + 2].strip('"').strip("'").strip()
    
    processor = Processor(docBakPath)
    if not processor.Open():
        print("Could not to open: " + docBakPath)
        continue
    
    isFound = False
    str = ""
    isRegexArg = inputArg.find("(") != -1 and inputArg.find("(") < inputArg.find(")")
    print(isRegexArg)
    if isRegexArg:
        isFound = processor.LocateRegexString(inputArg)
    else:
        isFound = processor.LocateString(inputArg)
    if not isFound:
        print("Could not find string: " + inputArg)
        continue

    print("Find: " + inputArg)
    if not processor.MarkString():
        print("Could not mark when handle: " + inputArg)
        continue
    
    if not processor.Save():
        print("Could not save when handle: " + inputArg)
        continue
    
    print("Handled: " + inputArg)
