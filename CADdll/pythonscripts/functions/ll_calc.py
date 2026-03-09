import clr

import acad
import academit

import System

def 命令(): 
    academit.添加命令("llcalc", llcalc)
    # academit.添加命令("ll-entsel", ll_entsel)

zhu_text_size = 50

@acad.decorator_command_undo
def llcalc():
    global zhu_text_size
    string = ""
    while True:
        char = acad.GetString("请输入表达式: ")
        if char == None: break
        string += char
    size = acad.GetInt(zhu_text_size, "请输入字体大小: ")
    if size > 0: zhu_text_size = size
    pt1  = acad.GetPoint("请点击放置位置: ")
    string = string.replace("=", "")
    string = string.replace("\n", "")
    if string != "":   
        result = eval(string)
        string = f"= {string} = {result}"
        acad.Prompt(string)
        with acad.transaction() as trans:
            acad.AddText(pt1, string, zhu_text_size)
    # print(string)


# def ll_entsel():
#     acad.GetActiveDocument()
#     acad.EntSel()