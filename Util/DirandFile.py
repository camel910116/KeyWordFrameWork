#encoding=utf-8
import os
from Util.FormatTime import *
from ProjectVar.Var import *

def createDir(path,dirName):
    dirPath = os.path.join(path,dirName)
    if os.path.exists(dirPath):
        pass
    else:
        os.mkdir(dirPath)

if __name__=="__main__":
    print screen_capturePicture_path + "\\"+dates()
    createDir(screen_capturePicture_path ,dates())