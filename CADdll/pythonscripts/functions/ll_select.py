


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
    academit.添加命令("llsl-pick-select", llsl_pick_select)
    academit.添加命令("llsl-fence", llsl_fence)
    academit.添加命令("llsl-fence-circle", llsl_fence_circle)
    academit.添加命令("llsl-entitysel", llsl_entitysel)
    academit.添加命令("llsl-entsel-textedit", llsl_entsel_textedit)
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
    acad.SSGet(sssetfirst=True)


@acad.decorator_command
def llsl_point(): 
    acad.SSGet([[0, "POINT"]], sssetfirst=True)

@acad.decorator_command
def llsl_line(): 
    acad.SSGet([[0, "LINE"]], sssetfirst=True)


@acad.decorator_command
def llsl_pl(): 
    flist = [
        [-4, "<OR"],[0, "LWPOLYLINE"],[0, "POLYLINE"], [-4, "OR>"]
        ]
    acad.SSGet(flist, sssetfirst=True)

@acad.decorator_command
def llsl_lwpl(): 
    acad.SSGet([[0, "LWPOLYLINE"]], sssetfirst=True)

@acad.decorator_command
def llsl_3dpl(): 
    acad.SSGet([[0, "POLYLINE"]], sssetfirst=True)

@acad.decorator_command
def llsl_spline(): 
    acad.SSGet([[0, "SPLINE"]], sssetfirst=True)


@acad.decorator_command
def llsl_arc(): 
    acad.SSGet([[0, "ARC"]], sssetfirst=True)


@acad.decorator_command
def llsl_circle(): 
    acad.SSGet([[0, "CIRCLE"]], sssetfirst=True)

@acad.decorator_command
def llsl_ellipse(): # 椭圆
    acad.SSGet([[0, "ELLIPSE"]], sssetfirst=True)

@acad.decorator_command
def llsl_region(): 
    acad.SSGet([[0, "REGION"]], sssetfirst=True)

@acad.decorator_command
def llsl_text(): 
    flist = [
        [-4, "<OR"],[0, "TEXT"],[0, "MTEXT"], [-4, "OR>"]
        ]
    acad.SSGet(flist, sssetfirst=True)

@acad.decorator_command
def llsl_dtext(): 
    acad.SSGet([[0, "TEXT"]], sssetfirst=True)

@acad.decorator_command
def llsl_mtext(): 
    acad.SSGet([[0, "MTEXT"]], sssetfirst=True)

@acad.decorator_command
def llsl_dim(): 
    flist = [
        [-4, "<OR"], [0, "DIMENSION"], [0, "*LEADER"], [-4, "OR>"]
        ]
    acad.SSGet(flist, sssetfirst=True)

@acad.decorator_command
def llsl_dimension(): 
    acad.SSGet([[0, "DIMENSION"]], sssetfirst=True)

@acad.decorator_command
def llsl_leader(): 
    acad.SSGet([[0, "LEADER"]], sssetfirst=True) # == qleader

@acad.decorator_command
def llsl_mleader(): 
    acad.SSGet([[0, "MULTILEADER"]], sssetfirst=True)

@acad.decorator_command
def llsl_block(): 
    flist = [
        [-4, "<OR"],[0, "*INSERT"], [2, "BLOCKDEFAULT*"], [-4, "OR>"]
        ]
    acad.SSGet(flist, sssetfirst=True)

@acad.decorator_command
def llsl_pick_select(): 
    pt1 = acad.GetPoint("请点击选择对象:")
    if pt1 == None: return 
    ss1 = acad.GetSelectCornerCross(pt1, pt1)
    acad.SSSetFirst(ss1)

@acad.decorator_command
def llsl_fence(): 
    pt1, pt2 = acad.GetPoint2()
    if pt1 == None: return 
    ss1 = acad.GetSelectFence(pt1, pt2)
    acad.SSSetFirst(ss1)

@acad.decorator_command
def llsl_fence_circle(): 
    pt1, pt2 = acad.GetPoint2()
    if pt1 == None: return 
    ss1 = acad.GetSelectFence(pt1, pt2, [[0, "CIRCLE"]])
    acad.SSSetFirst(ss1)

@acad.decorator_command
def llsl_entitysel(): 
    objid = acad.EntSel(string="请点击选择对象: ")  # EntSelEntity
    if acad.IsNoneObjectId(objid): return 
    ss1 = acad.SSSetFromIdList([objid])
    acad.SSSetFirst(ss1)


@acad.decorator_command
def llsl_entsel_textedit(): 
    objid = acad.EntSel(string="请点击选择对象: ")  # EntSelEntity
    if acad.IsNoneObjectId(objid): return 
    ss1 = acad.SSSetFromIdList([objid])
    acad.SSSetFirst(ss1)
    acad.Command(["TEXTEDIT"])







    
