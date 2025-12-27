import clr

import acad
import academit

import System

def 命令(): 
    academit.添加命令("ll-calc", ll_calc)
    academit.添加命令("ll-entsel", ll_entsel)

text_size = 50
def ll_calc():
    global text_size
    acad.GetActiveDocument()
    acad.GetUndo()
    try:
        string = ""
        while True:
            char = acad.GetString("请输入表达式: ")
            if char == "": break
            string += char
        size   = acad.GetInt(text_size, "请输入字体大小: ")
        if size > 0: text_size = size
        point  = acad.GetPoint()
        string = string.replace("=", "")
        string = string.replace("\n", "")
        if string != "": 
            result = eval(string)
            string = f"= {string} = {result}"
            acad.AddText(point, size, string)
        acad.SetUndo()
        print(string)
        print("END")
    except Exception as e:
        acad.HappenErrorUndo()
        print(e)
        acad.Prompt("...函数出错...\n")

    
def ll_entsel():
    acad.GetActiveDocument()
    acad.EntSel()