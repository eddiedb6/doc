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

    def Open(self):
        if not os.path.exists(self.__path):
            print("Path does not exist: " + self.__path)
            return False
        self.__doc = Document(self.__path)
        return True

    def FindString(self, string):
        isFount, locations = self.__locateString(string)
        print(isFount)
        print(locations)

    def __locateStringInRun(self, paragraph, paragraphIndex, beginIndex, endIndex):
        runIndex = 0
        runStartPos = 0
        runEndPos = 0
        runMapping = {}
        for run in paragraph.runs:
            runEndPos = runStartPos + len(run.text) - 1
            if runEndPos < beginIndex:
                # Not meet yet
                pass
            elif runStartPos > endIndex:
                # Over it already
                break
            elif runStartPos <= beginIndex and runEndPos >= beginIndex:
                # Start in this run
                if runEndPos >= endIndex:
                    # All belongs to this run
                    runMapping[runIndex] = [
                        beginIndex - runStartPos,
                        runEndPos - runStartPos,
                        True if beginIndex == runStartPos else False,
                        True if endIndex == runEndPos else False
                    ]
                    break
                else:
                    # Left part belongs to next run
                    runMapping[runIndex] = [
                        beginIndex - runStartPos,
                        runEndPos - runStartPos,
                        True if beginIndex == runStartPos else False,
                        False
                    ]
            elif runStartPos > beginIndex and runEndPos < endIndex:
                # This run is only middle part
                runMapping[runIndex] = [
                    0,
                    runEndPos - runStartPos,
                    True,
                    False
                ]
            elif runStartPos <= endIndex and runEndPos >= endIndex:
                # This run is end part
                runMapping[runIndex] = [
                    0,
                    endIndex - runStartPos,
                    True,
                    True if endIndex == runEndPos else False
                ]
                break
            else:
                print("Should not be here")
                return [False, None]
            runStartPos = runEndPos + 1
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
            print(paragraph.text)
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
        if self.__doc is None or string is None or string == "":
            return [False, None]
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
