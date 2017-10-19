import os.path

from docx import Document
from docx.shared import Inches

from DocLocation import Location

class Processor:
    def __init__(self, path):
        self.__path = path
        self.__doc = None

    def Open(self):
        if not os.path.exists(self.__path):
            print("Path does not exist: " + self.__path)
            return False
        self.__doc = Document(self.__path)
        return True

    def __locateString(self, string):
        if self.__doc is None or string is None or string == "":
            return False

        locations = []
        
        paragraphIndex = 0
        for paragraph in self.__doc.paragraphs:
            beginIndex = 0
            resultIndex = -1
            while True:
                resultIndex = paragraph.text.find(string, beginIndex)
                if resultIndex < 0:
                    break
                # One found in paragraph
                location = Location()
                location.ParagraphIndex = paragraphIndex
                
            paragraphIndex += 1
                
            
    

docPath = "Test.docx"
docModifyPath = "Test.md.docx"

docDebug = 1
parDebug = 1
runDebug = 0

needSave = 0

def Test():
    # Open doc
    doc = Document(docPath)
    if docDebug:
        print("[[Path]] " + docPath)
        print("[[Doc Objs]]")
        for obj in dir(doc):
            print("    " + obj)
            
    # Check each paragraph    
    lenParagraphs = len(doc.paragraphs)
    if parDebug:
        print("[[Paragraphs Len]] " + str(lenParagraphs))
    indexParagraph = 0
    for paragraph in doc.paragraphs:
        if parDebug:
            if indexParagraph == 0:
                print("[[Paragraph Objs]]")
                for obj in dir(paragraph):
                    print("    " + obj)
            print("[[Paragraph " + str(indexParagraph) + "]] " + paragraph.text)

        # Check each run
        lenRuns = len(paragraph.runs)
        if runDebug:
            print("    [[Runs Len]] " + str(lenRuns))
        indexRun = 0
        for run in paragraph.runs:
            if runDebug:
                if indexParagraph == 0 and indexRun == 0:
                    print("    [[Run Objs]]")
                    for obj in dir(run):
                        print("        " + obj)
                print("    [[Run " + str(indexRun) + "]] " + run.text)
            indexRun += 1

        indexParagraph += 1

    if needSave:
        doc.save(docModifyPath)

Test()
