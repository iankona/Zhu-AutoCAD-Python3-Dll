import os 
import sys
import importlib


dirpath = os.path.dirname(__file__)
# sys.path.append(dirpath+"\\libs") # 已经由dll加载
sys.path.append(dirpath+"\\ttks")
sys.path.append(dirpath+"\\libk")
sys.path.append(dirpath+"\\liby")
os.environ['TK_LIBRARY'] = dirpath+"\\ttks\\tcl\\tk8.6"
os.environ['TCL_LIBRARY'] = dirpath+"\\ttks\\tcl\\tcl8.6"


import clr
# from System import Console #, ConsoleColor

class NetConsole:
    def __init__(self):
        pass

    
    def write(self, message): # 1个print(), 调用2次write()
        message = str(message)
        # if message == '\n': return
        # if message[-1] != '\n': message += "\n"
        with open(r"F:\CADdll\pythonscripts\outputlog.txt", "a+") as log:
            log.write(message)

    def flush(self):
        pass
    
sys.stdout = NetConsole() # 接管print()输出 # print("你好，世界！")
sys.stderr = NetConsole()


# print(os.getcwd()) # C:\Users\Administrator\user\Documents
os.chdir("F:\\CADdll")

import coredumpy
coredumpy.patch_except(directory='F:\\CADdll\\pythonscripts\\dumps')
from coredumpy import config
# The dump depth if not specified
config.default_recursion_depth: int = 5

# Best effort timeout for dump - not guaranteed. Only checked for each new depth.
config.dump_timeout: int = 60

# Whether dump all threads
config.dump_all_threads: bool = False

# # Whether hide strings that match config.secret_patterns
# config.hide_secret: bool = True

# # The patterns for secrets
# config.secret_patterns: list[re.Pattern] = [re.compile(r"[A-Za-z0-9]{32,1024}")]

# # Whether hide strings that match os.environ.values()
# config.hide_environ: bool = True

# # The filter to determine whether an environ should be hidden
# config.environ_filter: Callable = lambda env: len(env) > 8



import academit


def 生成命令():
    名称列表 = []
    basenamelist = os.listdir(r"F:\CADdll\pythonscripts\functions")
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


