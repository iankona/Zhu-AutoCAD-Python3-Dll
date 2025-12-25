import os 
import sys
import importlib

import clr
# from System import Console #, ConsoleColor

class NetConsole:
    def __init__(self):
        pass
    
    def write(self, message): # 1个print(), 调用2次write()
        message = str(message)
        # if message == '\n': return
        # if message[-1] != '\n': message += "\n"
        with open(r"E:\CADdll\pythonscripts\outputlog.txt", "a+") as log:
            log.write(message)

    def flush(self):
        pass
    
sys.stdout = NetConsole() # 接管print()输出 # print("你好，世界！")
sys.stderr = NetConsole()


# print(os.getcwd()) # C:\Users\Administrator\user\Documents
os.chdir("E:\\CADdll")


import academit


def 生成命令():
    名称列表 = []
    basenamelist = os.listdir(r"E:\CADdll\pythonscripts\functions")
    for basename in basenamelist:
        if '__pycache__' in basename: continue
        charlist = basename.split('.')
        if charlist[0] == '': continue
        if charlist[0] not in 名称列表: 名称列表.append(charlist[0])

    for name in 名称列表: 
        try:
            module = importlib.import_module(name)
            module.命令()
            print(f"Python: '{name}' 加载成功 ... ")
        except Exception as e:
            print(f"Python: '{name}' 加载出错 ??? ")
            print(e)

    academit.保存程序集()



def 设置命令():
    academit.设置程序集()


