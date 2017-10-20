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
        for location in self.__locations:
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
        paragraph =  self.__doc.paragraphs[location.ParagraphIndex]
        if not self.__adjustRuns(paragraph, location):
            return False
        return True

    def __isStringStartInRun(self):
        return False

    def __locateStringInRun(self, paragraph, paragraphIndex, strBeginIndex, strEndIndex):
        runIndex = 0
        runBeginIndex = 0
        runEndIndex = 0
        runMapping = {}
        def isStringNotMetYet():
            return True if runEndIndex < strBeginIndex else False
        for run in paragraph.runs:
            runEndIndex = runBeginIndex + len(run.text) - 1
            if isStringNotMetYet():
                # Not meet yet
                continue
            elif runBeginIndex > strEndIndex:
                # Over it already
                break
            elif runBeginIndex <= strBeginIndex and runEndIndex >= strBeginIndex:
                # Start in this run
                if runEndIndex >= strEndIndex:
                    # All belongs to this run
                    runMapping[runIndex] = [
                        strBeginIndex - runBeginIndex,
                        runEndIndex - runBeginIndex,
                        True if strBeginIndex == runBeginIndex else False,
                        True if strEndIndex == runEndIndex else False
                    ]
                    break
                else:
                    # Left part belongs to next run
                    runMapping[runIndex] = [
                        strBeginIndex - runBeginIndex,
                        runEndIndex - runBeginIndex,
                        True if strBeginIndex == runBeginIndex else False,
                        True
                    ]
            elif runBeginIndex > strBeginIndex and runEndIndex < strEndIndex:
                # This run is only middle part
                runMapping[runIndex] = [
                    0,
                    runEndIndex - runBeginIndex,
                    True,
                    True
                ]
            elif runBeginIndex <= strEndIndex and runEndIndex >= strEndIndex:
                # This run is end part
                runMapping[runIndex] = [
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
        if len(runMapping) <= 0:
            return [False, None]
        location = Location()
        location.ParagraphIndex = paragraphIndex
        location.RunsCount = len(runMapping)
        i = 0
        for index in runMapping:
            location.SetRunIndex(i, index)
            location.SetStringRange(index, runMapping[index][0], runMapping[index][1], runMapping[index][2], runMapping[index][3])
            i += 1
        return [True, location]
        
    def __locateStringInParagraph(self, string, paragraph, paragraphIndex):
        locations = []
        beginIndex = 0
        while True:
            resultIndex = paragraph.text.find(string, beginIndex)
            if resultIndex < 0:
                break
            # One found in paragraph
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
