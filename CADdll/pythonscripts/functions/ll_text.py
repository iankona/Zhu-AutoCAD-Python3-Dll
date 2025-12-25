import clr

import acad
import academit

import System


def 命令(): 
    academit.添加命令("ll-text-songti", ll_text_songti)


text_size = 50
def ll_text_songti():
    acad.GetActiveDocument()
    acad.ChangeStandardFontStyle("宋体")

    
