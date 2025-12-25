import clr

import acad
import academit

import System

def 命令(): 
    # academit.添加命令("ll-calc", ll_calc)
    # academit.添加命令("ll-entsel", ll_entsel)
    pass

attach_length = 2.0
def ll_attach_line():
    global attach_length
    acad.GetActiveDocument()
    acad.GetUndo()
    try:
        pt1 = acad.GetPoint()
        string = string.replace("=", "")
        string = string.replace("\n", "")
        if string != "": 
            result = eval(string)
            string = f"= {string} = {result}"
            acad.AddText(point, size, string)
        acad.SetUndo()
    except Exception as e:
        acad.HappenErrorUndo()
        print(e)
        acad.Prompt("...函数出错...\n")

    
