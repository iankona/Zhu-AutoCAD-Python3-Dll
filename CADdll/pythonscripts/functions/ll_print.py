import clr

import acad
import academit

import System

def 命令(): 
    academit.添加命令("ll-print", ll_print)
    academit.添加命令("ll-print-bound", ll_print_bound)    
    academit.添加命令("ll-print-corner", ll_print_corner)
    academit.添加命令("ll-print-corner-cross", ll_print_corner_cross)
    academit.添加命令("ll-print-frence", ll_print_frence)
    academit.添加命令("ll-print-entsel", ll_print_entsel)
    academit.添加命令("ll-print-clayer", ll_print_current_layer)
    academit.添加命令("ll-print-point-in-bound", ll_print_point_in_bound)
    academit.添加命令("ll-print-include", ll_print_include)



@acad.decorator_command
def ll_print_include():
    objid1 = acad.EntSel()
    objid2 = acad.EntSel()
    acad.IsInclude(objid1, objid2)




@acad.decorator_command
def ll_print():
    objid = acad.EntSel()
    with acad.transaction() as trans:
        objref = acad.GetObjectForRead(objid)
        acad.Prompt(objref.GetType())



@acad.decorator_command
def ll_print_bound():
    objid = acad.EntSel()
    pt1, pt2 = acad.GetEntityBound(objid)
    acad.CommandAddPoint(pt1)
    acad.CommandAddPoint(pt2)



@acad.decorator_command
def ll_print_point_in_bound():
    objid = acad.EntSel()
    pt0 = acad.GetPoint()
    pt1, pt2 = acad.GetEntityBound(objid)
    acad.CommandAddPoint(pt0)
    acad.CommandAddPoint(pt1)
    acad.CommandAddPoint(pt2)
    flag  = acad.IsPointInRange(pt0, objid)
    acad.Prompt(flag)



@acad.decorator_command
def ll_print_entsel():
    ss1 =  acad.SSGet(sel_method=":S")
    acad.Prompt(ss1)



@acad.decorator_command
def ll_print_frence():
    pt1 = acad.GetPoint()
    pt2 = acad.GetPoint(base_point=pt1)
    # ss1 =  acad.SSGet(sel_method="+F")
    ss1 = acad.GetSelectFence(pt1, pt2)
    acad.Prompt(ss1)

@acad.decorator_command
def ll_print_corner_cross():
    pt1 = acad.GetPoint()
    pt2 = acad.GetCorner("", pt1)
    # ss1 =  acad.SSGet(sel_method="+F")
    ss1 = acad.GetSelectCornerCross(pt1, pt2)
    acad.Prompt(ss1)


@acad.decorator_command
def ll_print_corner():
    pt1 = acad.GetPoint()
    pt2 = acad.GetCorner("", pt1)
    # ss1 =  acad.SSGet(sel_method="+F")
    ss1 = acad.GetSelectCorner(pt1, pt2)
    acad.Prompt(ss1)



@acad.decorator_command
def ll_print_current_layer():
    with acad.transaction() as trans:
        objid = acad.db.Clayer
        layer_record = acad.GetObjectForRead(objid)
        acad.Prompt(layer_record.Name), acad.Prompt("\n")        
        acad.Prompt(layer_record.Color.ColorIndex)


