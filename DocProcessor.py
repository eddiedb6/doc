import os
import sys

from docx import Document
from docx.shared import Inches

sys.path.append(os.path.join(os.path.split(os.path.realpath(__file__))[0], "."))

from DocLocation import Location

class Processor:
    def __init__(self, path):
        self.__path = path
        self.__doc = None
        self.__locations = None

    def Open(self):
        if not os.path.exists(self.__path):
            print("[[DOC]] Path does not exist: " + self.__path)
            return False
        self.__doc = Document(self.__path)
        return True

    def FindString(self, string):
        if self.__doc is None:
            print("[[DOC]] Not ready for find yet")
            return False
        if string is None or string == "":
            print("[[DOC]] Invalid string for find")
            return False
        locateResult = self.__locateString(string)
        self.__locations = locateResult[1]
        return locateResult[0]

    def MarkString(self):
        if self.__doc is None:
            print("[[DOC]] Not ready for mark yet")
            return False
        if self.__locations is None:
            print("[[DOC]] No location for mark")
            return False
        locationsInParagraph = {}
        # Sort locations by paragraph and order by des
        for location in self.__locations:
            if location.ParagraphIndex in locationsInParagraph:
                locationsInParagraph[location.ParagraphIndex].insert(0, location)
            else:
                locationsInParagraph[location.ParagraphIndex] = [location]
        for paragraphIndex in locationsInParagraph:
            for location in locationsInParagraph[paragraphIndex]:
                if not self.__applyMark(location):
                    return False
        return True

    def Save(self):
        if self.__doc is None:
            print("[[DOC]] Not ready for save yet")
            return False
        self.__doc.save(self.__path)
        return True
        
    def __applyMark(self, location):
        paragraph = self.__doc.paragraphs[location.ParagraphIndex]
        if not self.__adjustRuns(paragraph, location):
            return False
        # TODO
        return True

    def __adjustRuns(self, paragraph, location):
        # Could only add but not 
        runsCount = location.RunsCount;
        runsToAdd = 0
        for i in range(0, runsCount):
            runIndex = location.GetRunIndex(i)
            stringRange = location.GetStringRange(runIndex)
            isFromBeginning = stringRange[2]
            isToEnd = stringRange[3]
            if not isFromBeginning:
                runsToAdd += 1
            if not isToEnd:
                runsToAdd += 1
        runsToAdd -= runsCount - 1
        while runsToAdd > 0:
            paragraph.add_run()
            runsToAdd -= 1
        return True

    def __locateStringInRun(self, paragraph, paragraphIndex, strBeginIndex, strEndIndex):
        runIndex = 0
        runBeginIndex = 0
        runEndIndex = 0
        strPosInRun = {}
        def isStringNotMetYet():
            return True if runEndIndex < strBeginIndex else False
        def isStringMetAlready():
            return True if runBeginIndex > strEndIndex else False
        def isStringBeginHere():
            return True if runBeginIndex <= strBeginIndex and runEndIndex >= strBeginIndex else False
        def isStringEndHere():
            return True if runBeginIndex <= strEndIndex and runEndIndex >= strEndIndex else False
        def isRunInMiddleOfString():
            return True if runBeginIndex > strBeginIndex and runEndIndex < strEndIndex else False
        for run in paragraph.runs:
            runEndIndex = runBeginIndex + len(run.text) - 1
            if isStringNotMetYet():
                pass
            elif isStringMetAlready():
                print("[[DOC]] Warning: should not come to here")
                break
            elif isStringBeginHere():
                print("[[DOC]] Start in run: " + str(runIndex))
                if isStringEndHere():
                    print("[[DOC]] Also end in run: " + str(runIndex))
                    strPosInRun[runIndex] = [
                        strBeginIndex - runBeginIndex,
                        runEndIndex - runBeginIndex,
                        True if strBeginIndex == runBeginIndex else False,
                        True if strEndIndex == runEndIndex else False
                    ]
                    break
                else:
                    # Left part belongs to next run
                    strPosInRun[runIndex] = [
                        strBeginIndex - runBeginIndex,
                        runEndIndex - runBeginIndex,
                        True if strBeginIndex == runBeginIndex else False,
                        True
                    ]
            elif isRunInMiddleOfString():
                print("[[DOC]] This run is in middle of string: " + str(runIndex))
                strPosInRun[runIndex] = [
                    0,
                    runEndIndex - runBeginIndex,
                    True,
                    True
                ]
            elif isStringEndHere():
                print("[[DOC]] End in run: " + str(runIndex))
                strPosInRun[runIndex] = [
                    0,
                    strEndIndex - runBeginIndex,
                    True,
                    True if strEndIndex == runEndIndex else False
                ]
                break
            else:
                print("[[DOC]] Should not be here")
                return [False, None]
            runBeginIndex = runEndIndex + 1
            runIndex += 1
        if len(strPosInRun) <= 0:
            return [False, None]
        location = Location()
        location.ParagraphIndex = paragraphIndex
        location.RunsCount = len(strPosInRun)
        count = 0
        for runIndex in strPosInRun:
            location.SetRunIndex(count, runIndex)
            location.SetStringRange(runIndex, strPosInRun[runIndex][0], strPosInRun[runIndex][1], strPosInRun[runIndex][2], strPosInRun[runIndex][3])
            count += 1
        return [True, location]
        
    def __locateStringInParagraph(self, string, paragraph, paragraphIndex):
        locations = []
        beginIndex = 0
        while True:
            resultIndex = paragraph.text.find(string, beginIndex)
            if resultIndex < 0:
                break
            # One found in paragraph
            print("[[DOC]] Find string in paragraph: " + str(paragraphIndex))
            locatedResult = self.__locateStringInRun(paragraph, paragraphIndex, resultIndex, len(string) - 1 + resultIndex)
            if locatedResult[0]:
                locations.append(locatedResult[1])
            beginIndex = resultIndex + 1
        if len(locations) > 0:
            return [True, locations]
        return [False, None]

    def __locateString(self, string):
        locations = []
        paragraphIndex = 0
        for paragraph in self.__doc.paragraphs:
            locateResult = self.__locateStringInParagraph(string, paragraph, paragraphIndex)
            if locateResult[0]:
                for location in locateResult[1]:
                    locations.append(location)
            paragraphIndex += 1
        if len(locations) > 0:
            return [True, locations]
        return [False, None]
