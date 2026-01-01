import clr

import acad
import academit

import System

def 命令(): 
    academit.添加命令("ll-calc", ll_calc)
    # academit.添加命令("ll-entsel", ll_entsel)

text_size = 50

@acad.decorator_command_undo
def ll_calc():
    global text_size
    string = ""
    while True:
        char = acad.GetString("请输入表达式: ")
        if char == None: break
        string += char
    size = acad.GetInt(text_size, "请输入字体大小: ")
    if size > 0: text_size = size
    pt1  = acad.GetPoint("请点击放置位置: ")
    string = string.replace("=", "")
    string = string.replace("\n", "")
    if string != "": 
        result = eval(string)
        string = f"= {string} = {result}"
        acad.CommandAddText(pt1, string, text_size)
    # print(string)


# def ll_entsel():
#     acad.GetActiveDocument()
#     acad.EntSel()