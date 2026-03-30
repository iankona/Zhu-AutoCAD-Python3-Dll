
from Autodesk.AutoCAD.DatabaseServices import DBPoint, Extents3d, Polyline, Polyline3d, Line, Circle, Poly3dType, DBText, Region, DBObjectCollection, Intersect


import acad
import academit
import System

def 命令(): 
    academit.添加命令("llbbrn", llbbrn)
    academit.添加命令("llbbrn-for", llbbrn_for)
    academit.添加命令("llbbrn-auto-for", llbbrn_auto_for)
    academit.添加命令("llbbrn-offset", llbbrn_offset)
    academit.添加命令("llbbrn-offset-for", llbbrn_offset_for)
    academit.添加命令("llbbrn-offset-auto-for", llbbrn_offset_auto_for)





@acad.decorator_command
def llbbrn():
    objidlist =  acad.SSGetIdList()   
    pt1, pt2 = acad.GetIdListBoundXY0(objidlist)
    with acad.transaction() as trans:
        acad.AddRect(pt1, pt2)

@acad.decorator_command
def llbbrn_for():
    while True:
        pt1, pt2 = acad.GetCorner2("请选择第1个顶点: ", "请选择第2个顶点: ")
        if pt1 == None: return
        objidlist = acad.GetSelectCornerIdList(pt1, pt2)
        pt1, pt2 = acad.GetIdListBoundXY0(objidlist)
        with acad.transaction() as trans:
            acad.AddRect(pt1, pt2)


@acad.decorator_command
def llbbrn_auto_for():
    color_index = acad.CountColorIndex()
    objidlist = acad.SSGetIdList()  
    with acad.transaction() as trans:
        result = acad.TransAutoFindRegionRectList(objidlist)
        for pt1, pt2 in result: 
            acad.AddRect(pt1, pt2, color_index=color_index)


zhu_bbrn_offset = 50
@acad.decorator_command
def llbbrn_offset():
    global zhu_bbrn_offset
    length = acad.GetDouble(zhu_bbrn_offset, "请输入外偏移大小:")
    if length != None: zhu_bbrn_offset = length
    objidlist =  acad.SSGetIdList()   
    pt1, pt2 = acad.GetIdListBoundXY0(objidlist)
    with acad.transaction() as trans:
        objref1 = acad.DBObjectRect(pt1, pt2)
        collect = objref1.GetOffsetCurves(zhu_bbrn_offset)
        for objref2 in collect: acad.AddDBObject(objref2)



@acad.decorator_command
def llbbrn_offset_for():
    global zhu_bbrn_offset
    length = acad.GetDouble(zhu_bbrn_offset, "请输入外偏移大小:")
    if length != None: zhu_bbrn_offset = length
    while True:
        pt1, pt2 = acad.GetCorner2("请选择第1个顶点: ", "请选择第2个顶点: ")
        if pt1 == None: break
        objidlist = acad.GetSelectCornerIdList(pt1, pt2)
        pt1, pt2 = acad.GetIdListBoundXY0(objidlist)
        with acad.transaction() as trans:
            objref1 = acad.DBObjectRect(pt1, pt2)
            collect = objref1.GetOffsetCurves(zhu_bbrn_offset)
            for objref2 in collect: acad.AddDBObject(objref2)


@acad.decorator_command
def llbbrn_offset_auto_for():
    global zhu_bbrn_offset
    length = acad.GetDouble(zhu_bbrn_offset, "请输入外偏移大小:")
    if length != None: zhu_bbrn_offset = length
    color_index = acad.CountColorIndex()
    objidlist = acad.SSGetIdList()  
    with acad.transaction() as trans:
        result = acad.TransAutoFindRegionRectList(objidlist)
        for pt1, pt2 in result: 
            objref1 = acad.DBObjectRect(pt1, pt2)
            collect = objref1.GetOffsetCurves(zhu_bbrn_offset)
            for objref2 in collect: 
                objref2.ColorIndex = color_index
                acad.AddDBObject(objref2)








# if acad.IsCCW(ptnlist):
#     collect = objref1.GetOffsetCurves( zhu_square_size)  
# else:
#     collect = objref1.GetOffsetCurves(-zhu_square_size) 