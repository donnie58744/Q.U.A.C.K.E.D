import os
import cv2
from gekko import GEKKO
class calcBitrate():
    def getClipSize(self, filename):
        b = os.path.getsize(filename)
        return b

    def getClipDuration(self, filename):
        # create video capture object
        data = cv2.VideoCapture(rf'{filename}')
        # count the number of frames
        frames = data.get(cv2.CAP_PROP_FRAME_COUNT)
        fps = data.get(cv2.CAP_PROP_FPS)
        
        # calculate duration of the video
        seconds = frames / fps
        return seconds
    def solveEquation(self, filename):
        seconds = self.getClipDuration(filename)
        base = 7.2
        if (seconds >= 40):
            base = 6.5
        finished = (((base*8)/seconds)*1000)
        return round(finished)