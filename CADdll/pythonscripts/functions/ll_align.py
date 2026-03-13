
from Autodesk.AutoCAD.DatabaseServices import DBPoint, Extents3d, Polyline, Polyline3d, Line, Circle, Poly3dType, DBText, Region, DBObjectCollection, Intersect



import acad
import academit
import System

def 命令(): 
    academit.添加命令("llalign-pmin-x", llalign_pmin_x)
    academit.添加命令("llalign-pmin-y", llalign_pmin_y)
    academit.添加命令("llalign-pmax-x", llalign_pmax_x)
    academit.添加命令("llalign-pmax-y", llalign_pmax_y)
    academit.添加命令("llalign-center-x", llalign_center_x)
    academit.添加命令("llalign-center-y", llalign_center_y)



@acad.decorator_command
def llalign_pmin_x():
    objidlist = acad.SSGetIdList(string="请选择对齐对象: ")  
    pb1 = acad.GetPoint("请选择对齐基点: ")
    if pb1 == None: return
    xb, yb, zb = pb1
    with acad.transaction() as trans:
        result = acad.TransAutoFindRegionRectList(objidlist)
        for pt1, pt2 in result: 
            x1, y1, z1 = pt1
            po1 = [0, y1, 0]
            po2 = [0, yb, 0]
            objidlist = acad.GetSelectCornerCrossIdList(pt1, pt2)
            for objid in objidlist: acad.TransMove(objid, po1, po2)

            

@acad.decorator_command
def llalign_pmin_y():
    objidlist = acad.SSGetIdList(string="请选择对齐对象: ")  
    pb1 = acad.GetPoint("请选择对齐基点: ")
    if pb1 == None: return
    xb, yb, zb = pb1
    with acad.transaction() as trans:
        result = acad.TransAutoFindRegionRectList(objidlist)
        for pt1, pt2 in result: 
            x1, y1, z1 = pt1
            po1 = [x1, 0, 0]
            po2 = [xb, 0, 0]
            objidlist = acad.GetSelectCornerCrossIdList(pt1, pt2)
            for objid in objidlist: acad.TransMove(objid, po1, po2)



@acad.decorator_command
def llalign_pmax_x():
    objidlist = acad.SSGetIdList(string="请选择对齐对象: ")  
    pb1 = acad.GetPoint("请选择对齐基点: ")
    if pb1 == None: return
    xb, yb, zb = pb1
    with acad.transaction() as trans:
        result = acad.TransAutoFindRegionRectList(objidlist)
        for pt1, pt2 in result: 
            x1, y1, z1 = pt2
            po1 = [0, y1, 0]
            po2 = [0, yb, 0]
            objidlist = acad.GetSelectCornerCrossIdList(pt1, pt2)
            for objid in objidlist: acad.TransMove(objid, po1, po2)

            

@acad.decorator_command
def llalign_pmax_y():
    objidlist = acad.SSGetIdList(string="请选择对齐对象: ")  
    pb1 = acad.GetPoint("请选择对齐基点: ")
    if pb1 == None: return
    xb, yb, zb = pb1
    with acad.transaction() as trans:
        result = acad.TransAutoFindRegionRectList(objidlist)
        for pt1, pt2 in result: 
            x1, y1, z1 = pt2
            po1 = [x1, 0, 0]
            po2 = [xb, 0, 0]
            objidlist = acad.GetSelectCornerCrossIdList(pt1, pt2)
            for objid in objidlist: acad.TransMove(objid, po1, po2)


@acad.decorator_command
def llalign_center_x():
    objidlist = acad.SSGetIdList(string="请选择对齐对象: ")  
    pb1 = acad.GetPoint("请选择对齐基点: ")
    if pb1 == None: return
    xb, yb, zb = pb1
    with acad.transaction() as trans:
        result = acad.TransAutoFindRegionRectList(objidlist)
        for pt1, pt2 in result: 
            x1, y1, z1 = [(pt1[0]+pt2[0])/2, (pt1[1]+pt2[1])/2, 0]
            po1 = [0, y1, 0]
            po2 = [0, yb, 0]
            objidlist = acad.GetSelectCornerCrossIdList(pt1, pt2)
            for objid in objidlist: acad.TransMove(objid, po1, po2)

            

@acad.decorator_command
def llalign_center_y():
    objidlist = acad.SSGetIdList(string="请选择对齐对象: ")  
    pb1 = acad.GetPoint("请选择对齐基点: ")
    if pb1 == None: return
    xb, yb, zb = pb1
    with acad.transaction() as trans:
        result = acad.TransAutoFindRegionRectList(objidlist)
        for pt1, pt2 in result: 
            x1, y1, z1 = [(pt1[0]+pt2[0])/2, (pt1[1]+pt2[1])/2, 0]
            po1 = [x1, 0, 0]
            po2 = [xb, 0, 0]
            objidlist = acad.GetSelectCornerCrossIdList(pt1, pt2)
            for objid in objidlist: acad.TransMove(objid, po1, po2)

