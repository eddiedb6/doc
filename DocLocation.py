class Location:
    def __init__(self):
        self.__paragraphIndex = -1
        self.__runsIndex = []
        self.__stringMap = {}

    @property
    def ParagraphIndex(self):
        return self.__paragraphIndex

    @ParagraphIndex.setter
    def ParagraphIndex(self. value):
        self.__paragraphIndex = value
    
    @property
    def RunsCount(self):
        return len(self.__runsIndex)

    @RunsCount.setter
    def RunsCount(self. value):
        if value <= 0:
            return
        self.__runsIndex = []
        self.__stringMap = {}
        for i in range(0, value):
            self.__runsIndex.append(-1)

    def GetRunIndex(self, index):
        if index < 0 || index >= len(self.__runsIndex):
            return -1
        return self.__runsIndex[index]

    def SetRunIndex(self, index, value):
        if index < 0 || index >= len(self.__runsIndex):
            return
        self.__stringMap[value] = [-1, -1, False, False]
        return self.__runsIndex[index] = value

    def GetStringRange(self, runIndex):
        if runIndex in self.__stringMap:
            return self.__stringMap[runIndex]
        return [-1, -1, False, False]

    def SetStringRange(self, runIndex, start, end, isFromBeginning, isToEnd):
        self.__stringMap[runIndex] = [start, end, isFromBeginning, isToEnd]
        
        
