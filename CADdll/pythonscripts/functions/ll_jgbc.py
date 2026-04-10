import clr

import acad
import academit

import System
import clr
import random
import acad
import academit
from Autodesk.AutoCAD.DatabaseServices import SubentityId, SubentityType, FullSubentityPath, MeshFaceterData, IdMapping, ObjectIdCollection, SubDMesh, Curve, Extents3d, Polyline, Polyline3d, Line, Circle, Poly3dType, DBText, Region, BooleanOperationType
from Autodesk.AutoCAD.Geometry import Point2d, Point3d, Point3dCollection, Matrix3d, Vector3d, NurbCurve3d
import System
from Autodesk.AutoCAD.BoundaryRepresentation import PointContainment, Brep, Face, BrepEntity
import Autodesk
from Autodesk.AutoCAD.Colors import Color, ColorMethod
from Autodesk.AutoCAD.DatabaseServices import Entity, DBPoint, Extents3d, Polyline, Polyline3d, Line, Circle, Poly3dType, DBText, Region, DBObjectCollection, Intersect

def 命令(): 
    academit.添加命令("lljgbc", lljgbc)
    # academit.添加命令("lljgbc-sheet-for", lljgbc_sheet_for)




llzhu_jgbc_inside_width = 0.5
llzhu_jgbc_outside_width = 0.5
def llzhu_ui_jiguanbuchang_input():
    global llzhu_jgbc_inside_width, llzhu_jgbc_outside_width
    ilength = acad.GetDouble(llzhu_jgbc_inside_width, "请输入内线补偿:")
    olength = acad.GetDouble(llzhu_jgbc_outside_width, "请输入外线补偿:")
    if ilength != None: llzhu_jgbc_inside_width = ilength
    if olength != None: llzhu_jgbc_outside_width = olength



def llzhu_trans_jiguanbuchang_add_region_polyline(objidlist):
    collect = DBObjectCollection()
    for objid in objidlist:
        objref = acad.TransObjectForRead(objid)
        if objref.Layer == "0": collect.Add(objref)
    regions = Region.CreateFromCurves(collect)
    for objref in regions:
        resultlist = acad.DBObjectConvertRegionToPolylineXY0(objref)
        for polyline in resultlist: acad.AddDBObject(polyline, "补偿1", 3)   




def llzhu_trans_jiguanbuchang_add_region_spline(objidlist):
    collect = DBObjectCollection()
    for objid in objidlist:
        objref = acad.TransObjectForRead(objid)
        if objref.Layer == "0": collect.Add(objref)
    regions = Region.CreateFromCurves(collect)

             


    buflist = []
    for objref in regions: buflist.append([objref.Area, objref])
    buflist.sort(key = lambda item: item[0], reverse=True) # 从大到小
    
    resultlist = []
    for region in regions:
        brep = Brep(region)
        for face in brep.Faces: 
            for loop in face.Loops:
                reflist = []
                for edge in loop.Edges:
                    nc3d = edge.GetCurveAsNurb()
                    dbcurve = Curve.CreateFromGeCurve(nc3d)
                    reflist.append(dbcurve)
                resultlist.append(reflist)

    for reflist in resultlist:
        bufline = reflist[0]
        for objref in reflist[1:]:
            bufline.JoinEntity(objref)  # 生成的Spline
        acad.AddDBObject(bufline, "补偿1", 3)   



def llzhu_trans_jiguanbuchang_auto_offset(objidlist):
    collect = DBObjectCollection()
    for objid in objidlist:
        objref = acad.TransObjectForWrite(objid)
        collect.Add(objref)
    regions = Region.CreateFromCurves(collect)





@acad.decorator_command
def lljgbc():
    llzhu_ui_jiguanbuchang_input()
    objidlist = acad.SSGetIdList()   
    pt1, pt2 = acad.GetPoint2("请选择基点: ", "请选择目标位置")
    with acad.transaction() as trans:
        objidlist = acad.TransAutoExplodeObjectIdList(objidlist)
        result = acad.TransAutoFindRegionRectList(objidlist)
    with acad.transaction() as trans: # 涉及前置对象需要先提交cad
        for pd1, pd2 in result: 
            objidlist = acad.GetSelectCornerCrossIdList(pd1, pd2)
            # llzhu_trans_jiguanbuchang_add_region(objidlist)

    # objidlist = acad.GetSelectCornerCrossIdList(po1, po2, [[0, "REGION"]])
    # ss1 = acad.SSSetFromIdList(objidlist)
    # acad.Command(["move", ss1, "", acad.ToPoint3d(pt1), acad.ToPoint3d(pt2)]), acad.Prompt("\n")
    # llzhu_trans_jiguanbuchang_auto_offset(objidlist)



# @acad.decorator_command
# def lljgbc_sheet_for():
#     llzhu_ui_jiguanbuchang_input()
#     objidlist = acad.SSGetIdList()   
#     with acad.transaction() as trans:
#         result = llzhu_trans_jiguanbuchang_add_region_polyline(objidlist)
        # for pline in result:
        #     acad.AddDBObject(pline, "补偿1", 3)   


# 核心函数JoinEntities和JoinEntity成功运行，但转换出来的多段线曲线无法偏移


    # collect = DBObjectCollection()
    # for objid in objidlist:
    #     objref = acad.TransObjectForRead(objid)
    #     if objref.Layer == "0": collect.Add(objref)
    # regions = Region.CreateFromCurves(collect)

             


    # buflist = []
    # for objref in regions: buflist.append([objref.Area, objref])
    # buflist.sort(key = lambda item: item[0], reverse=True) # 从大到小
    
    # resultlist = []
    # for region in regions:
    #     brep = Brep(region)
    #     for face in brep.Faces: 
    #         for loop in face.Loops:
    #             reflist = []
    #             for edge in loop.Edges:
    #                 nc3d = edge.GetCurveAsNurb()
    #                 dbcurve = Curve.CreateFromGeCurve(nc3d)
    #                 reflist.append(dbcurve)
    #             resultlist.append(reflist)

    # for reflist in resultlist:
    #     bufline = reflist[0]
    #     for objref in reflist[1:]:
    #         bufline.JoinEntity(objref) 
    #     acad.AddDBObject(bufline, "补偿1", 3)   







#     错误，无法保证前后对象是相连的，导致核心函数JoinEntities和JoinEntity运行错误
#     collect = DBObjectCollection()
#     for objid in objidlist:
#         objref = acad.TransObjectForRead(objid)
#         if objref.Layer == "0": collect.Add(objref)
#     regions = Region.CreateFromCurves(collect)
#     # for objref in regions: 
#     #     acad.CheckLayerAndColor(objref, "面域1", 35)
#     #     acad.currentblock.AppendEntity(objref)
#     #     acad.trans.AddNewlyCreatedDBObject(objref, True) # 不用先提交到CAD，可以Explode
#     resultlist = []
#     for region in regions:
#         result = DBObjectCollection()
#         region.Explode(result)
#         for objref in result: 
#             acad.CheckLayerAndColor(objref, "补偿1", 35)
#             acad.currentblock.AppendEntity(objref)
#             acad.trans.AddNewlyCreatedDBObject(objref, True) # 不用先提交到CAD，可以Explode               
#         resultlist.append(result)

#     buflist = []
#     for objref in regions: buflist.append([objref.Area, objref])
#     buflist.sort(key = lambda item: item[0], reverse=True) # 从大到小
    
#     for result in resultlist:
#         bufline = result[0]
#         entlist = System.Array[Entity](result.Count-1) # 'Entity[]' 
#         for i in range(result.Count-1):
#             # objref = acad.DBObjectConvertLineToPolyline(objref)
#             entlist[i] = result[i+1]
#             # acad.CheckLayerAndColor(objref, "补偿1", 35)
#             # acad.currentblock.AppendEntity(objref)
#             # acad.trans.AddNewlyCreatedDBObject(objref, True) # 不用先提交到CAD，可以Explode 
#         bufline.JoinEntities(entlist)
#         # acad.AddDBObject(pline, "补偿1", 3)   

# # # 相联结的曲线不要求是同种类型的曲线，例如，直线与圆弧也可以联结成一条曲线，但是，曲线之间必须是连续的。
# # .JoinEntities(Entity[]) # 要求曲线是连续的
# # .JoinEntity(Entity)     # 要求曲线是连续的, 