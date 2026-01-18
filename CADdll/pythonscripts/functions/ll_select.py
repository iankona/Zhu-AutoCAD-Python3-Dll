


import clr

import acad
import academit

import System

def 命令(): 
    academit.添加命令("llsl-color", llsl_color)
    academit.添加命令("llsl-layer", llsl_layer)
    academit.添加命令("llsl-layer-and-color", llsl_layer_and_color)
    # academit.添加命令("llsl-color-and-layer", llsl_layer_and_color)
    academit.添加命令("llsl", llsl)
    academit.添加命令("llsl-point", llsl_point)
    academit.添加命令("llsl-line", llsl_line)
    academit.添加命令("llsl-lwpl", llsl_lwpl)
    academit.添加命令("llsl-3dpl", llsl_3dpl)
    academit.添加命令("llsl-spline", llsl_spline)
    academit.添加命令("llsl-arc", llsl_arc)
    academit.添加命令("llsl-circle", llsl_circle)
    academit.添加命令("llsl-ellipse", llsl_ellipse)
    academit.添加命令("llsl-region", llsl_region)
    academit.添加命令("llsl-text", llsl_text)
    academit.添加命令("llsl-dtext", llsl_dtext)
    academit.添加命令("llsl-mtext", llsl_mtext)
    academit.添加命令("llsl-dim", llsl_dim)
    academit.添加命令("llsl-dimension", llsl_dimension)
    academit.添加命令("llsl-leader", llsl_leader)
    academit.添加命令("llsl-mleader", llsl_mleader)
    academit.添加命令("llsl-block", llsl_block)
    pass


@acad.decorator_command
def llsl_color(): 
    objlist1 = acad.SSGetIdList()
    color_index_list1 = acad.GetIdListColorIndexList(objlist1)
    acad.SSSetFirst(None)
    objlist2 = acad.SSGetIdList()
    color_index_list2 = acad.GetIdListColorIndexList(objlist2)
    acad.SSSetFirst(None)
    objidlist3 = []
    for objid, index in zip(objlist2, color_index_list2):
        if index in color_index_list1: objidlist3.append(objid)
        # if index in color_index_list1 and objid not in objidlist3: objidlist3.append(objid)
    ss2 = acad.SSSetFromIdList(objidlist3)
    acad.SSSetFirst(ss2)


@acad.decorator_command
def llsl_layer(): 
    objlist1 = acad.SSGetIdList()
    layer_name_list1 = acad.GetIdListLayerNameList(objlist1)
    acad.SSSetFirst(None)
    objlist2 = acad.SSGetIdList()
    layer_name_list2 = acad.GetIdListLayerNameList(objlist2)
    acad.SSSetFirst(None)
    objidlist3 = []
    for objid, name in zip(objlist2, layer_name_list2):
        if name in layer_name_list1: objidlist3.append(objid)
    ss2 = acad.SSSetFromIdList(objidlist3)
    acad.SSSetFirst(ss2)


@acad.decorator_command
def llsl_layer_and_color(): 
    objlist1 = acad.SSGetIdList()
    layer_name_list1 = acad.GetIdListLayerNameList(objlist1)
    color_index_list1 = acad.GetIdListColorIndexList(objlist1)
    acad.SSSetFirst(None)
    objlist2 = acad.SSGetIdList()
    layer_name_list2 = acad.GetIdListLayerNameList(objlist2)
    color_index_list2 = acad.GetIdListColorIndexList(objlist2)
    acad.SSSetFirst(None)
    objidlist3 = []
    for objid, name, index in zip(objlist2, layer_name_list2, color_index_list2):
        if name in layer_name_list1 and index in color_index_list1: objidlist3.append(objid)
    ss2 = acad.SSSetFromIdList(objidlist3)
    acad.SSSetFirst(ss2)


@acad.decorator_command
def llsl(): 
    acad.SSGet()


@acad.decorator_command
def llsl_point(): 
    acad.SSGet([[0, "POINT"]])

@acad.decorator_command
def llsl_line(): 
    acad.SSGet([[0, "LINE"]])


@acad.decorator_command
def llsl_pl(): 
    flist = [
        [-4, "<OR"],[0, "LWPOLYLINE"],[0, "POLYLINE"], [-4, "OR>"]
        ]
    acad.SSGet(flist)

@acad.decorator_command
def llsl_lwpl(): 
    acad.SSGet([[0, "LWPOLYLINE"]])

@acad.decorator_command
def llsl_3dpl(): 
    acad.SSGet([[0, "POLYLINE"]])

@acad.decorator_command
def llsl_spline(): 
    acad.SSGet([[0, "SPLINE"]])


@acad.decorator_command
def llsl_arc(): 
    acad.SSGet([[0, "ARC"]])


@acad.decorator_command
def llsl_circle(): 
    acad.SSGet([[0, "CIRCLE"]])

@acad.decorator_command
def llsl_ellipse(): # 椭圆
    acad.SSGet([[0, "ELLIPSE"]])

@acad.decorator_command
def llsl_region(): 
    acad.SSGet([[0, "REGION"]])

@acad.decorator_command
def llsl_text(): 
    flist = [
        [-4, "<OR"],[0, "TEXT"],[0, "MTEXT"], [-4, "OR>"]
        ]
    acad.SSGet(flist)

@acad.decorator_command
def llsl_dtext(): 
    acad.SSGet([[0, "TEXT"]])

@acad.decorator_command
def llsl_mtext(): 
    acad.SSGet([[0, "MTEXT"]])

@acad.decorator_command
def llsl_dim(): 
    flist = [
        [-4, "<OR"], [0, "DIMENSION"], [0, "*LEADER"], [-4, "OR>"]
        ]
    acad.SSGet(flist)

@acad.decorator_command
def llsl_dimension(): 
    acad.SSGet([[0, "DIMENSION"]])

@acad.decorator_command
def llsl_leader(): 
    acad.SSGet([[0, "LEADER"]]) # == qleader

@acad.decorator_command
def llsl_mleader(): 
    acad.SSGet([[0, "MULTILEADER"]])

@acad.decorator_command
def llsl_block(): 
    flist = [
        [-4, "<OR"],[0, "*INSERT"], [2, "BLOCKDEFAULT*"], [-4, "OR>"]
        ]
    acad.SSGet(flist)