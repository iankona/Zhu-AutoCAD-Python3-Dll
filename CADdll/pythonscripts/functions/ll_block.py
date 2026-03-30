
from Autodesk.AutoCAD.DatabaseServices import DBPoint, Extents3d, Polyline, Polyline3d, Line, Circle, Poly3dType, DBText, Region, DBObjectCollection, Intersect


import acad
import academit
import System

def 命令(): 
    academit.添加命令("llblock", llblock)
    academit.添加命令("llblock-for", llblock_for)
    academit.添加命令("llblock-auto-for", llblock_auto_for)
    academit.添加命令("llblock-offset", llblock_offset)
    academit.添加命令("llblock-offset-for", llblock_offset_for)
    academit.添加命令("llblock-offset-auto-for", llblock_offset_auto_for)

    academit.添加命令("llexplode", llexplode)



llzhu_align_filter = [
    [-4, "<OR"], [0, "LINE"], [0, "LWPOLYLINE"], [0, "SPLINE"],  [0, "ARC"], [0, "CIRCLE"], [0, "ELLIPSE"], [-4, "OR>"]
    ]

llzhu_align_filter = [[0, "LWPOLYLINE"]]
llzhu_align_filter = []


@acad.decorator_command
def llexplode():
    objidlist =  acad.SSGetIdList()   
    with acad.transaction() as trans:
        for objid in objidlist:
            result_objidlist = acad.TransExplode(objid)
            # print(result_objidlist)


@acad.decorator_command
def llblock():
    objidlist =  acad.SSGetIdList()   
    pt1, pt2 = acad.GetIdListBoundXY0(objidlist)
    with acad.transaction() as trans:
        acad.AddBlockFromIdList(objidlist, pt1)

@acad.decorator_command
def llblock_for():
    while True:
        pt1, pt2 = acad.GetCorner2("请选择第1个顶点: ", "请选择第2个顶点: ")
        if pt1 == None: break
        objidlist = acad.GetSelectCornerIdList(pt1, pt2)
        pt1, pt2 = acad.GetIdListBoundXY0(objidlist)
        with acad.transaction() as trans:
            acad.AddBlockFromIdList(objidlist, pt1)



@acad.decorator_command
def llblock_auto_for():
    objidlist = acad.SSGetIdList(llzhu_align_filter)  
    with acad.transaction() as trans:
        result = acad.TransAutoFindRegionRectList(objidlist)
        for pt1, pt2 in result: 
            objidlist = acad.GetSelectCornerCrossIdList(pt1, pt2)
            acad.AddBlockFromIdList(objidlist, pt1)







zhu_block_offset = 50
@acad.decorator_command
def llblock_offset():
    global zhu_block_offset
    length = acad.GetDouble(zhu_block_offset, "请输入外偏移大小:")
    if length != None: zhu_block_offset = length
    objidlist =  acad.SSGetIdList()   
    pt1, pt2 = acad.GetIdListBoundXY0(objidlist)
    with acad.transaction() as trans:
        objref1 = acad.DBObjectRect(pt1, pt2)
        collect = objref1.GetOffsetCurves(zhu_block_offset)
        extend = Extents3d()
        for objref2 in collect: 
            extend.AddExtents(objref2.GeometricExtents)
        pt1, pt2 = [extend.MinPoint.X, extend.MinPoint.Y, extend.MinPoint.Z], [extend.MaxPoint.X, extend.MaxPoint.Y, extend.MaxPoint.Z]
        objidlist = acad.GetSelectCornerCrossIdList(pt1, pt2)
        acad.AddBlockFromIdList(objidlist, pt1)

@acad.decorator_command
def llblock_offset_for():
    global zhu_block_offset
    length = acad.GetDouble(zhu_block_offset, "请输入外偏移大小:")
    if length != None: zhu_block_offset = length
    while True:
        pt1, pt2 = acad.GetCorner2("请选择第1个顶点: ", "请选择第2个顶点: ")
        if pt1 == None: break
        objidlist = acad.GetSelectCornerIdList(pt1, pt2)
        pt1, pt2 = acad.GetIdListBoundXY0(objidlist)
        with acad.transaction() as trans:
            objref1 = acad.DBObjectRect(pt1, pt2)
            collect = objref1.GetOffsetCurves(zhu_block_offset)
            extend = Extents3d()
            for objref2 in collect: 
                extend.AddExtents(objref2.GeometricExtents)
            pt1, pt2 = [extend.MinPoint.X, extend.MinPoint.Y, extend.MinPoint.Z], [extend.MaxPoint.X, extend.MaxPoint.Y, extend.MaxPoint.Z]
            objidlist = acad.GetSelectCornerCrossIdList(pt1, pt2)
            acad.AddBlockFromIdList(objidlist, pt1)





@acad.decorator_command
def llblock_offset_auto_for():
    global zhu_block_offset
    length = acad.GetDouble(zhu_block_offset, "请输入外偏移大小:")
    if length != None: zhu_block_offset = length
    objidlist = acad.SSGetIdList(llzhu_align_filter)  
    with acad.transaction() as trans:
        result = acad.TransAutoFindRegionRectList(objidlist)
        for pt1, pt2 in result: 
            objref1 = acad.DBObjectRect(pt1, pt2)
            collect = objref1.GetOffsetCurves(zhu_block_offset)
            extend = Extents3d()
            for objref2 in collect: 
                extend.AddExtents(objref2.GeometricExtents)
            pt1, pt2 = [extend.MinPoint.X, extend.MinPoint.Y, extend.MinPoint.Z], [extend.MaxPoint.X, extend.MaxPoint.Y, extend.MaxPoint.Z]
            objidlist = acad.GetSelectCornerCrossIdList(pt1, pt2)
            acad.AddBlockFromIdList(objidlist, pt1)








# if acad.IsCCW(ptnlist):
#     collect = objref1.GetOffsetCurves( zhu_square_size)  
# else:
#     collect = objref1.GetOffsetCurves(-zhu_square_size) 