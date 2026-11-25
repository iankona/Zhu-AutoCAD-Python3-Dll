import os 
import sys

dirpath = os.path.dirname(__file__)
sys.path.append(dirpath+"\\libs") 
sys.path.append(dirpath+"\\ttks")
sys.path.append(dirpath+"\\libk")
sys.path.append(dirpath+"\\liby")
os.environ['TK_LIBRARY'] = dirpath+"\\ttks\\tcl\\tk8.6"
os.environ['TCL_LIBRARY'] = dirpath+"\\ttks\\tcl\\tcl8.6"


sys.path.append(dirpath+"\\functions")

workdir = "E:\\CADdll"
acaddir = "D:\\Program Files\\Autodesk\\AutoCAD 2023"
zhuddll = "Zhu.Acad.Plugin"