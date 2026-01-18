import clr

import acad
import academit

import System


def 命令(): 
    academit.添加命令("ll-text-songti", ll_text_songti)


text_size = 50

@acad.decorator_command
def ll_text_songti():
    acad.ChangeStandardFontStyle("宋体")

    
