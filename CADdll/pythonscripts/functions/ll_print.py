import clr

import acad
import academit

import System

def 命令(): 
    academit.添加命令("ll-print", ll_print)
    academit.添加命令("ll-print-clayer", ll_print_current_layer)

@acad.decorator_command
def ll_print():
    objid = acad.EntSel()
    with acad.transaction() as trans:
        objref = acad.GetObjectForRead(objid)
        acad.Prompt(objref.GetType())

@acad.decorator_command
def ll_print_current_layer():
    with acad.transaction() as trans:
        objid = acad.db.Clayer
        layer_record = acad.GetObjectForRead(objid)
        acad.Prompt(layer_record.Name), acad.Prompt("\n")        
        acad.Prompt(layer_record.Color.ColorIndex)


