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
from Autodesk.AutoCAD.DatabaseServices import DBPoint, Extents3d, Polyline, Polyline3d, Line, Circle, Poly3dType, DBText, Region, DBObjectCollection, Intersect

def 命令(): 
    academit.添加命令("llregion", llregion)
    academit.添加命令("llregion-cut", llregion_cut)
    # academit.添加命令("llregion-sheet", llregion_sheet)
    academit.添加命令("llregion-alignX", llregion_alignX)
    academit.添加命令("llregion-alignY", llregion_alignY)
    # academit.添加命令("llregion-subtract", llregion_subtract)
    # academit.添加命令("llregion-print", llregion_print)
    academit.添加命令("llregion-normal", llregion_normal)
    # academit.添加命令("llregion-check-inside", llregion_check_inside)
    # academit.添加命令("llregion-mesh-sphm", llregion_mesh_sphm)
    # academit.添加命令("llregion-objidlist", llregion_objidlist) 
    academit.添加命令("llregion-rotate", llregion_rotate) 
    academit.添加命令("llregion-flatten", llregion_flatten)
    academit.添加命令("llregion-rotate-to-z-up", llregion_rotate_to_z_up)
    academit.添加命令("llregion-ss-rotate-to-z-up", llregion_ss_rotate_to_z_up)
    academit.添加命令("llregion-align-xy0-point4", llregion_align_xy0_point4)
    academit.添加命令("llregion-align-point4", llregion_align_point4)
    academit.添加命令("llregion-ss-face-rotate-to-z-up", llregion_ss_face_rotate_to_z_up)
    # academit.添加命令("llregion-rotate-point-cloud-to-z-up", llregion_rotate_point_cloud_to_z_up)


# 可以通过检查ed.CurrentUserCoordinateSystem是否为单位矩阵(Matrix3d.Identity) 来判断
# 如果ed.CurrentUserCoordinateSystem == Matrix3d.Identity, 则UCS与WCS一致, 无需转换
# Matrix.Inverse()

llzhu_sheet_ness = 1.0
def llzhu_ui_region_sheet_ness():
    global llzhu_sheet_ness
    length = acad.GetDouble(llzhu_sheet_ness, "请输入板厚:")
    if length != None: llzhu_sheet_ness = length



def llzhu_region_sheet_auto_cut(objidlist):
    collect = DBObjectCollection()
    for objid in objidlist:
        objref = acad.TransObjectForWrite(objid)
        collect.Add(objref)
    regions = Region.CreateFromCurves(collect)
    for objref in regions: acad.AddDBObject(objref, "面域1", 35)

    area_dbobj_list = []
    for objref in regions: area_dbobj_list.append([objref.Area, objref])
    area_dbobj_list.sort(key = lambda item: item[0], reverse=True) # 数值差距大，导致python调用了C#的排序出错，使用lambda使运行正常，原理不明
    check_index_list = [0]
    for i, [area, objref] in enumerate(area_dbobj_list):
        if i in check_index_list: continue
        check_index_list.append(i)
        large_objref = objref
        child_objref_list = []
        for j, [area, objref] in enumerate(area_dbobj_list[i+1:]):
            extend = objref.GeometricExtents
            center = Point3d( (extend.MinPoint.X+extend.MaxPoint.X)/2, (extend.MinPoint.Y+extend.MaxPoint.Y)/2, (extend.MinPoint.Z+extend.MaxPoint.Z)/2 )
            brep = Brep(large_objref)
            result = brep.GetPointContainment(center, PointContainment.Outside) # out ref 被自动处理成结果返回
            brep_entity = result[0]
            if brep_entity == None: continue
            if str(brep_entity.GetType()) == "Autodesk.AutoCAD.BoundaryRepresentation.Face": # 点在面内
                check_index_list.append(i+1+j)
                child_objref_list.append(objref)
        for objref in child_objref_list: large_objref.BooleanOperation(BooleanOperationType.BoolSubtract, objref)
    area_dbobj_list[0][1].Erase(True)


@acad.decorator_command
def llregion():
    llzhu_ui_region_sheet_ness()
    objidlist = acad.SSGetIdList()   
    pt1, pt2 = acad.GetPoint2("请选择基点: ", "请选择目标位置")
    with acad.transaction() as trans:
        po1, po2 = acad.TransObjectIdListBoundXY0(objidlist)
        objidlist = acad.TransExplodeIdListWithErase(objidlist)
        result = acad.TransAutoFindRegionRectList(objidlist)
    with acad.transaction() as trans: # 涉及前置对象需要先提交cad
        for pd1, pd2 in result: 
            objidlist = acad.GetSelectCornerCrossIdList(pd1, pd2)
            llzhu_region_sheet_auto_cut(objidlist)


    objidlist = acad.GetSelectCornerCrossIdList(po1, po2, [[0, "REGION"]])
    ss1 = acad.SSSetFromIdList(objidlist)
    acad.Command(["move", ss1, "", acad.ToPoint3d(pt1), acad.ToPoint3d(pt2)]), acad.Prompt("\n")
    acad.Command(["EXTRUDE", ss1, "", System.Double(llzhu_sheet_ness)])
       
 




def __trans_ptlist_from_region(objid):
    objref = acad.TransObjectForWrite(objid)
    brep = acad.Brep(objref)
    ptlist = []
    for vertex in brep.Vertices:
        pt1 = [vertex.Point.X, vertex.Point.Y, vertex.Point.Z]
        ptlist.append(pt1)
    return ptlist

def find_same_point(ptlist1, ptlist2):
    ptlist0 = []
    for pt1 in ptlist1:
        for pt2 in ptlist2: 
            if acad.IsPointSame(pt1, pt2): ptlist0.append(pt1)
    return ptlist0

def find_connet_list(buflist):
    objid0 = buflist[0][0][0]
    result = []
    sklist = []
    for i in range(len(buflist)):
        for values in buflist:
            for objid1, objid2 in values:
                if objid0 == objid1: 
                    if objid2 in sklist: continue # 防止回环
                    result.append([objid1, objid2])
                    objid0 = objid2
                    sklist.append(objid1)
    return result

@acad.decorator_command
def llregion_flatten():
    objidlist = acad.SSGetIdList([[0, "REGION"]])
    objid0 = acad.EntSel([[0, "REGION"]], string="请选择基准对象:") 
    objidlist.remove(objid0)    
    with acad.transaction() as trans:
        baselist0 = __trans_ptlist_from_region(objid0)
        bufptlist = [[objid0, baselist0]]
        for objid in objidlist:
            ptlist = __trans_ptlist_from_region(objid)
            bufptlist.append([objid, ptlist])

    buflist = []
    for objid1, ptlist1 in bufptlist:
        cuflist = []
        for objid2, ptlist2 in bufptlist:
            if objid1 == objid2: continue
            pslist = find_same_point(ptlist1, ptlist2)
            if pslist == []: continue
            cuflist.append([objid1, objid2])
        buflist.append(cuflist)

    if len(buflist[0]) == 1:
        buflist1 = find_connet_list(buflist)
        buflist2 = None
    if len(buflist[0]) == 2:
        a = buflist[0][0]
        b = buflist[0][1]
        buflist1 = find_connet_list([[a]]+buflist[1:])
        buflist2 = find_connet_list([[b]]+buflist[1:])

    with acad.transaction() as trans:
        __trans_flatten_(buflist1)
    if buflist2 == None: return 
    with acad.transaction() as trans:
        __trans_flatten_(buflist2)

def __trans_flatten_(buflist):
    for i, [objid1, objid2] in enumerate(buflist):
        ptlist1 = __trans_ptlist_from_region(objid1)
        ptlist2 = __trans_ptlist_from_region(objid2)
        pslist1 = find_same_point(ptlist1, ptlist2)
        objref1 = acad.TransObjectForWrite(objid1)
        objref2 = acad.TransObjectForWrite(objid2)
        dr1 = [objref1.Normal.X, objref1.Normal.Y, objref1.Normal.Z]
        dr2 = [objref2.Normal.X, objref2.Normal.Y, objref2.Normal.Z]
        axis = acad.CrossNormalized(dr1, dr2)
        angle = acad.AngleFromDotDr1Dr2(dr1, dr2)
        acad.TransRoation(objid2, -angle, axis, pslist1[0])
        for objid3, objid4 in buflist[i+1:]:
            acad.TransRoation(objid4, -angle, axis, pslist1[0])


    

# zhu_objid_dict = {}
# @acad.decorator_command
# def llregion_objidlist():
#     global zhu_objid_dict
#     pt1 = acad.GetPoint()
#     pt2 = acad.GetPoint(base_point=pt1)
#     objidlist = acad.GetSelectFenceIdList(pt1, pt2, [[0, "REGION"]])
#     objidlist = [objid for objid in objidlist]
#     zhu_objid_dict = {}
#     for i, objid in enumerate(objidlist):
#         key = str(objid)
#         zhu_objid_dict[key] = objidlist[i:]


# @acad.decorator_command
# def llregion_rotate():
#     objid1 = acad.EntSel()
#     if acad.IsNone(objid1): return
#     pt1 = acad.GetPoint()
#     pt2 = acad.GetPoint()
#     angle = acad.GetDouble(zhu_angle, "请输入旋转角度:")
#     if pt1 == None or pt2 == None  or angle == None: return
#     zhu_angle = angle
#     axis = acad.Direct(pt1, pt2)
#     key = str(objid1)
#     if key not in zhu_objid_dict:
#         acad.Prompt("选择的对象未在列表里, 请先栏选生成列表...")
#         return
#     objidlist = zhu_objid_dict[key]
#     with acad.transaction() as trans:
#         for objid in objidlist:
#             acad.TransRoation(objid, zhu_angle, axis, pt1)




@acad.decorator_command
def llregion_rotate():
    objidlist = acad.SSGetIdList()
    objid0 = acad.EntSel(string="请选择基准对象:")
    objidlist.remove(objid0)
    while True:
        objid1 = acad.EntSel(string="请选择旋转对象:")
        if acad.IsNone(objid1): return 
        po1, po2 = acad.GetPoint2("请选择旋转轴第1点:", "请选择旋转轴第2点:")
        if po1 == None: return 
        pt1, pt3 = acad.GetPoint2()
        if pt1 == None: return 
        normal = acad.Direct(po1, po2)
        normal = acad.Vec3ResetLength(normal, 1)
        pt1 = acad.PointToPointNormalPlane(pt1, po2, normal)
        pt3 = acad.PointToPointNormalPlane(pt3, po2, normal)
        # pt1, po2，pt3
        dr2 = acad.Direct(pt1, po2)
        dr1 = acad.Direct(po2, pt3)
        dr1 = acad.Vec3ResetLength(dr1, 10)
        dr2 = acad.Vec3ResetLength(dr2, 10)
        normal = acad.CrossNormalized(dr1, dr2)
        angle = acad.AngleFromDotDr1Dr2(dr1, dr2)
        with acad.transaction() as trans:
            acad.TransRoationIdList(objidlist, angle, normal, po2)
        objidlist.remove(objid1)
        objid0 = objid1





@acad.decorator_command
def llregion_normal():
    objidlist = acad.SSGetIdList([[0, "REGION"]])
    with acad.transaction() as trans:
        for i, objid in enumerate(objidlist):
            objref = acad.TransObjectForRead(objid)
            extend = objref.GeometricExtents
            point1 = extend.MinPoint
            point2 = extend.MaxPoint
            center = [(point1.X+point2.X)/2, (point1.Y+point2.Y)/2, (point1.Z+point2.Z)/2]
            normal = objref.Normal
            length = acad.Distance([point1.X, point1.Y, point1.Z], [point2.X, point2.Y, point2.Z])
            direct = acad.Vec3ResetLength([normal.X, normal.Y, normal.Z], length/2)
            acad.AddLine(center, acad.Vec3Add(center, direct))
            acad.AddLine(center, acad.Vec3Add(center, [0,0,length/2]))
            angle = acad.AngleFromDotDr1Dr2([normal.X, normal.Y, normal.Z], [0, 0, 1])
            acad.Prompt(angle)



# def llzhu_trans_region_rotate_to_z_up(pt1):
#     objid1 = acad.GetSelectPickId(pt1)
#     normal = acad.GetEntityNormal(objid1)
#     normal = [normal.X, normal.Y, normal.Z]
#     angle1 = acad.AngleFromDotDr1Dr2(normal, [0,0, 1])
#     angle2 = acad.AngleFromDotDr1Dr2(normal, [0,0,-1])
#     if angle1 >= angle2:
#         axis = acad.CrossNormalized(normal, [0, 0, -1])
#         angle = angle2
#     else:
#         axis = acad.CrossNormalized(normal, [0, 0,  1])
#         angle = angle1
#     acad.TransRoation(objid1, angle, axis, pt1)



@acad.decorator_command
def llregion_alignX():
    objidlist = acad.SSGetIdList() 
    buflist = []
    for objid in objidlist:
        pt1, pt2 = acad.GetEntityBoundXY0(objid)
        x1, y1, z1 = pt1
        x2, y2, z2 = pt2 
        center = [(x1+x2)/2, (y1+y2)/2, (z1+z2)/2]
        length = x2-x1
        width = y2-y1
        buflist.append([x1, 0, objid, pt1, pt2, length, width])
    buflist.sort(key = lambda item: item[0], reverse=False) # 小 -> 大

    x0 = buflist[0][0]
    for i, [x1, dx, objid, pt1, pt2, length, width] in enumerate(buflist):
        buflist[i][0] = x0
        buflist[i][1] = x0 - x1 # dx
        x0 = x0 + length + 10
    with acad.transaction() as trans:
        for x1, dx, objid1, pt1, pt2, length, width in buflist:
            po1 = acad.Vec3Add(pt1, [dx, 0, 0])
            acad.TransMove(objid1, pt1, po1)

@acad.decorator_command
def llregion_alignY():
    objidlist = acad.SSGetIdList() 
    buflist = []
    for objid in objidlist:
        pt1, pt2 = acad.GetEntityBoundXYZ(objid)
        x1, y1, z1 = pt1
        x2, y2, z2 = pt2 
        center = [(x1+x2)/2, (y1+y2)/2, (z1+z2)/2]
        length = x2-x1
        width = y2-y1
        buflist.append([y1, 0, objid, pt1, pt2, length, width])
    buflist.sort(key = lambda item: item[0], reverse=False) # 小 -> 大

    y0 = buflist[0][0]
    for i, [y1, dy, objid, pt1, pt2, length, width] in enumerate(buflist):
        buflist[i][0] = y0
        buflist[i][1] = y0 - y1 # dx
        y0 = y0 + width + 10
    with acad.transaction() as trans:
        for y1, dy, objid1, pt1, pt2, length, width in buflist:
            po1 = acad.Vec3Add(pt1, [0, dy, 0])
            acad.TransMove(objid1, pt1, po1)



@acad.decorator_command
def llregion_rotate_to_z_up():
    while True:
        pt1 = acad.GetPoint()
        if pt1 == None: return
        objid1 = acad.GetSelectPickId(pt1, [[0, "REGION"]])
        normal = acad.GetEntityNormal(objid1)
        normal = [normal.X, normal.Y, normal.Z]
        angle1 = acad.AngleFromDotDr1Dr2(normal, [0,0, 1])
        angle2 = acad.AngleFromDotDr1Dr2(normal, [0,0,-1])
        if angle1 >= angle2:
            axis = acad.CrossNormalized(normal, [0, 0, -1])
            angle = angle2
        else:
            axis = acad.CrossNormalized(normal, [0, 0,  1])
            angle = angle1
        with acad.transaction() as trans:
            acad.TransRoation(objid1, angle, axis, pt1)
        acad.Prompt(angle), acad.Prompt("\n")



@acad.decorator_command
def llregion_ss_rotate_to_z_up():
    while True:
        objidlist = acad.SSGetIdList([[0, "REGION"]])
        objid1 = acad.EntSel([[0, "REGION"]], string="请选择基准对象:") 
        if acad.IsNone(objid1): return
        pt1 = acad.GetPoint("请选择基点: ")
        if pt1 == None: return
        normal = acad.GetEntityNormal(objid1)
        normal = [normal.X, normal.Y, normal.Z]
        angle1 = acad.AngleFromDotDr1Dr2(normal, [0,0, 1])
        angle2 = acad.AngleFromDotDr1Dr2(normal, [0,0,-1])
        if angle1 >= angle2:
            axis = acad.CrossNormalized(normal, [0, 0, -1])
            angle = angle2
        else:
            axis = acad.CrossNormalized(normal, [0, 0,  1])
            angle = angle1
        with acad.transaction() as trans:
            for objid1 in objidlist:
                acad.TransRoation(objid1, angle, axis, pt1)
        acad.Prompt(angle), acad.Prompt("\n")



@acad.decorator_command
def llregion_ss_face_rotate_to_z_up():
    while True:
        objidlist = acad.SSGetIdList()
        pt1, pt2, pt3 = acad.GetPoint3() 
        if pt1 == None: return
        dr1 = acad.Direct(pt1, pt2)
        dr2 = acad.Direct(pt2, pt3)
        normal = acad.CrossNormalized(dr1, dr2)
        angle1 = acad.AngleFromDotDr1Dr2(normal, [0,0, 1])
        angle2 = acad.AngleFromDotDr1Dr2(normal, [0,0,-1])
        if angle1 >= angle2:
            axis = acad.CrossNormalized(normal, [0, 0, -1])
            angle = angle2
        else:
            axis = acad.CrossNormalized(normal, [0, 0,  1])
            angle = angle1
        with acad.transaction() as trans:
            for objid1 in objidlist:
                acad.TransRoation(objid1, angle, axis, pt1)
        acad.Prompt(angle), acad.Prompt("\n")








@acad.decorator_command
def llregion_align_xy0_point4():
    while True:
        objid1 = acad.EntSel()
        pt1 = acad.GetPoint()
        # objid1 = acad.GetSelectPickId(pt1)
        normal = acad.GetEntityNormal(objid1)
        normal = [normal.X, normal.Y, normal.Z]
        angle1 = acad.AngleFromDotDr1Dr2(normal, [0,0, 1])
        angle2 = acad.AngleFromDotDr1Dr2(normal, [0,0,-1])
        if angle1 >= angle2:
            axis = acad.CrossNormalized(normal, [0, 0, -1])
            angle = angle2
        else:
            axis = acad.CrossNormalized(normal, [0, 0,  1])
            angle = angle1
        with acad.transaction() as trans:
            acad.TransRoation(objid1, angle, axis, pt1)
        pt2, po2, po1 = acad.GetPoint3("第2点", "第3点", "第4点",)
        dr1 = acad.Direct(pt2[0:2], pt1[0:2])
        dr2 = acad.Direct(po2[0:2], po1[0:2])
        angle = acad.AngleFromDotDr1Dr2(dr1, dr2)
        axis = acad.CrossNormalized(dr1, dr2) # 如果 axis = [0,0,1] 还是需要用Cross判断是外扩方向，还是内缩方向，还是判断angle正负
        with acad.transaction() as trans:
            acad.TransRoation(objid1, angle, axis, pt2)
            acad.TransMove(objid1, pt2, po2)


@acad.decorator_command
def llregion_align_point4():
    while True:
        objidlist = acad.SSGetIdList()
        if objidlist == None: return
        pt1, pt2 = acad.GetPoint2()
        po2, po1 = acad.GetPoint2()
        dr1 = acad.Direct(pt2, pt1)
        dr2 = acad.Direct(po2, po1)
        angle = acad.AngleFromDotDr1Dr2(dr1, dr2)
        axis = acad.CrossNormalized(dr1, dr2)
        with acad.transaction() as trans:
            for objid in objidlist:
                acad.TransRoation(objid, angle, axis, pt2)
                acad.TransMove(objid, pt2, po2)


# @acad.decorator_command
# def llregion_rotate_point_cloud_to_z_up():
#     objid1 = acad.EntSel([[0, "REGION"]], string="请选择参考面:")
#     pt1 = acad.GetPoint()
#     pt2 = acad.GetPoint()
#     normal = acad.GetEntityNormal(objid1)
#     normal = [normal.X, normal.Y, normal.Z]
#     angle = acad.AngleFromDotDr1Dr2(normal, [0,0,1])
#     acad.Prompt(angle)
#     axis = acad.CrossNormalized(normal, [0,0,1])
#     # with acad.command_undo():
#     #     direct = acad.Direct(pt1, pt2)
#     #     po1, po2 = pt1, pt2
#     #     if acad.Dot(axis, direct) < 0: po1, po2 = pt2, pt1
#     #     acad.CommandRotate3d(objid1, po1, po2, angle)
#     objidlist = acad.SSGetIdList(string="请选择需要旋转的对象:")
#     with acad.transaction() as trans:
#         for objid in objidlist:
#             acad.TransRoation(objid, angle, axis, pt1)
#     # acad.Prompt(angle)




   
# @acad.decorator_command
# def llregion_sheet():
#     dbobj_collect = acad.SSGetDbObjectCollection()
#     with acad.transaction() as trans:
#         regions = acad.AddRegion(dbobj_collect, "面域1", 35)
#         area_dbobj_list = []
#         for objref in regions: area_dbobj_list.append([objref.Area, objref])
#         area_dbobj_list.sort(key = lambda item: item[0], reverse=True) # 数值差距大，导致python调用了C#的排序出错，使用lambda使运行正常，原理不明
#         check_index_list = [0]
#         for i, [area, objref] in enumerate(area_dbobj_list):
#             if i in check_index_list: continue
#             check_index_list.append(i)
#             large_objref = objref
#             child_objref_list = []
#             for j, [area, objref] in enumerate(area_dbobj_list[i+1:]):
#                 extend = objref.GeometricExtents
#                 center = Point3d( (extend.MinPoint.X+extend.MaxPoint.X)/2, (extend.MinPoint.Y+extend.MaxPoint.Y)/2, (extend.MinPoint.Z+extend.MaxPoint.Z)/2 )
#                 brep = Brep(large_objref)
#                 result = brep.GetPointContainment(center, PointContainment.Outside) # out ref 被自动处理成结果返回
#                 brep_entity = result[0]
#                 if brep_entity == None: continue
#                 if str(brep_entity.GetType()) == "Autodesk.AutoCAD.BoundaryRepresentation.Face": # 点在面内
#                     check_index_list.append(i+1+j)
#                     child_objref_list.append(objref)
#             for objref in child_objref_list: large_objref.BooleanOperation(BooleanOperationType.BoolSubtract, objref)
#         area_dbobj_list[0][1].Erase(True)
# # (<Autodesk.AutoCAD.BoundaryRepresentation.Face object at 0x00000227DE697380>, <PointContainment.OnBoundary: 2>) Autodesk.AutoCAD.BoundaryRepresentation.Face
# # (None, <PointContainment.Outside: 1>)




@acad.decorator_command
def llregion_cut():
    dbobj_collect = acad.SSGetDbObjectCollection()
    with acad.transaction() as trans:
        regions = acad.AddRegion(dbobj_collect, "面域1", 35)
        max_area, max_objref = 0, None
        dbobj_list = []
        for objref in regions:
        # for objid in objidlist:
        #     objref = acad.TransObjectForWrite(objid)
            dbobj_list.append(objref)
            if objref.Area > max_area:
                max_area = objref.Area
                max_objref = objref
        # print(max_area, max_objref)
        if max_objref != None:
            # dbobj_list.remove(max_objref)
            index = dbobj_list.index(max_objref)
            obj_list = dbobj_list[0:index] + dbobj_list[index+1:]
            for objref in obj_list: max_objref.BooleanOperation(BooleanOperationType.BoolSubtract, objref)


# @acad.decorator_command
# def llregion_check_inside():
#     objid1 = acad.EntSel(string="请选择第1个对象:")
#     objid2 = acad.EntSel(string="请选择第2个对象:")
#     with acad.transaction() as trans:
#         region1 = acad.TransObjectForWrite(objid1)
#         region2 = acad.TransObjectForWrite(objid2)
#         extend = region2.GeometricExtents
#         point1 = extend.MinPoint
#         point2 = extend.MaxPoint # [point1.X, point1.Y, point1.Z], [point2.X, point2.Y, point2.Z]         
#         center_point = Point3d((point1.X+point2.X)/2, (point1.Y+point2.Y)/2, (point1.Z+point2.Z)/2)
#         brep = Brep(region1)
#         result = brep.GetPointContainment(center_point, PointContainment.Outside) # out ref 被自动处理成结果返回
#         brep_entity = result[0]
#         # print(brep_entity.GetType())
#         # print(ent) # (<Autodesk.AutoCAD.BoundaryRepresentation.Face object at 0x00000204409347C0>, <PointContainment.OnBoundary: 2>)
#         if str(brep_entity.GetType()) == "Autodesk.AutoCAD.BoundaryRepresentation.Face":
#             acad.Prompt("点在面范围里")
#             # region1.BooleanOperation(BooleanOperationType.BoolSubtract, region2)
            
# @acad.decorator_command
# def llregion_subtract():
#     objidlist = acad.SSGetIdList([[0, "REGION"]])
#     with acad.transaction() as trans:
#         max_area, max_objref = 0, None
#         dbobj_list = []
#         for objid in objidlist:
#             objref = acad.TransObjectForWrite(objid)
#             dbobj_list.append(objref)
#             if objref.Area > max_area:
#                 max_area = objref.Area
#                 max_objref = objref
#         # print(max_area, max_objref)
#         if max_objref != None:
#             # dbobj_list.remove(max_objref)
#             index = dbobj_list.index(max_objref)
#             obj_list = dbobj_list[0:index] + dbobj_list[index+1:]
#             for objref in obj_list: max_objref.BooleanOperation(BooleanOperationType.BoolSubtract, objref)



# @acad.decorator_command
# def llregion_print():
#     objidlist = acad.SSGetIdList()
#     with acad.transaction() as trans:
#         for objid in objidlist:
#             objref = acad.TransObjectForWrite(objid)
#             normal = objref.Normal # (0.554700196225229,0,0.832050294337844) 
#             x,y,z = objref.Normal
#             matrix4x41 = objref.Ecs # ((1,0,0,0),(0,1,0,0),(0,0,1,0),(0,0,0,1))
#             # matrix4x42 = Matrix3d.PlaneToWorld(Vector3d(0,0,1))
#             # matrix4x43 = Matrix3d.PlaneToWorld(objref.Normal)
#             # matrix4x44 = Matrix3d.PlaneToWorld(Vector3d(-x,0,1-z))
#             # print(objref, objref.Normal)
#             # objref.TransformBy(matrix4x44)
#             # print(objref, objref.Normal)



# @acad.decorator_command
# def llregion_mesh_sphm():
#     with acad.transaction() as trans:
#         pass
#         # SphericalMesh
#         mySubD = SubDMesh()
#         mySubD.SetSphere(20, 8, 8, 0)
#         mySubD.SetDatabaseDefaults()
#         acad.AddDBObject(mySubD)

#         # cloneAndInflateSolid
#         solidId = mySubD.ObjectId
#         frac = 0.2
#         ids = ObjectIdCollection()
#         ids.Add(solidId)
#         im = IdMapping()
#         acad.db.DeepCloneObjects(ids, acad.db.CurrentSpaceId, im, False)
#         newId = im.Lookup(solidId).Value
#         sol = acad.TransObjectForWrite(newId)
#         ext = sol.Bounds
#         vec = ext.MaxPoint - ext.MinPoint
#         # # sol.OffsetBody(vec.Length * frac)
#         # fd = MeshFaceterData(0.01 * vec.Length, 40 * System.Math.PI / 180, 2, 2, 15, 5, 5, 0 )
#         # mdc = SubDMesh.GetObjectMesh(sol, fd)
#         # sd = SubDMesh()
#         # sd.SetDatabaseDefaults()
#         # sd.SetSubDMesh(mdc.VertexArray, mdc.FaceArray, 1)
#         # # sol.Erase()

#         # meshId = newId
#         # creasePaths =System.List[FullSubentityPath]
#         # sd.UpgradeOpen()
#         # subEnts = sd.GetSubentities()
#         # for subEnt in subEnts:
#         #     if subEnt.FullSubentityPath.SubentId.Type == SubentityType.Edge:
#         #         creasePaths.Add(subEnt.FullSubentityPath)
 
#         # sd.SetCrease(creasePaths.ToArray(), -1)

#         # bumpPaths =System.List[FullSubentityPath]
#         # sd.UpgradeOpen()
#         # subEnts = sd.GetSubentities()

#         # extrudelen = 1.0
#         # rnd = System.Random()
 
#         # for subEnt in subEnts:
#         #     if subEnt.FullSubentityPath.SubentId.Type == SubentityType.Face:
#         #         faces = FullSubentityPath[1]
#         #         faces[0] = subEnt.FullSubentityPath
#         #         fPlane = sd.GetFacePlane(faces[0].SubentId)
#         #         norm = fPlane.Normal
#         #         ext = sd.Bounds
#         #         vec = ext.MaxPoint - ext.MinPoint
#         #         mat = Matrix3d.Displacement(norm * 0.5 * extrudelen *  rnd.NextDouble() * vec.Length)
#         #         sd.TransformSubentityPathsBy(faces, mat)

#         # rnd = System.Random()
#         # sd.UpgradeOpen()
#         # for i in range(sd.NumberOfFaces - 1):
#         #     sId = SubentityId(SubentityType.Face, i)
#         #     r,g,b = random.randint(1, 255),random.randint(1, 255),random.randint(1, 255)
#         #     col = Color.FromRgb(r,g,b)
#         #     sd.SetSubentColor(sId, col)

#         # sd.UpgradeOpen()
#         # sol = sd.ConvertToSolid(True, True)
#         # newId = acad.AddDBObject(sol)

 
#         # ext = sol.Bounds
#         # pt1 = Point3d(ext.MaxPoint.X, ext.MinPoint.Y, ext.MinPoint.Z)
#         # transVec = 1.5 * (pt1 - ext.MinPoint)
#         # mat = Matrix3d.Displacement(transVec)
#         # sol.TransformBy(mat)






# using Autodesk.AutoCAD.Geometry;

# Point3d pointA = new Point3d(1, 2, 3);
# Point3d pointB = new Point3d(4, 5, 6);
# Point3d pointC = new Point3d(7, 8, 9);

# // 计算向量
# Vector3d vectorAB = pointB - pointA;
# Vector3d vectorAC = pointC - pointA;

# // 点乘
# double dotProduct = vectorAB.DotProduct(vectorAC);

# // 叉乘
# Vector3d crossProduct = vectorAB.CrossProduct(vectorAC);

# // 输出结果
# Console.WriteLine($"点乘结果: {dotProduct}");
# Console.WriteLine($"叉乘结果: {crossProduct}");


# @acad.decorator_command
# def llregion_rotate_to_z_up_negative():
#     objid1 = acad.EntSel([[0, "REGION"]])
#     pt1 = acad.GetPoint()
#     pt2 = acad.GetPoint()
#     normal = acad.GetEntityNormal(objid1)
#     normal = [normal.X, normal.Y, normal.Z]
#     angle = acad.AngleFromDotDr1Dr2(normal, [0,0,-1])
#     axis = acad.CrossNormalized(normal, [0,0,-1])
#     # with acad.command_undo():
#     #     direct = acad.Direct(pt1, pt2)
#     #     po1, po2 = pt1, pt2
#     #     if acad.Dot(axis, direct) < 0: po1, po2 = pt2, pt1
#     #     acad.CommandRotate3d(objid1, po1, po2, angle)
#     with acad.transaction() as trans:
#         acad.TransRoation(objid1, angle, axis, pt1)
#     acad.Prompt(angle)


# import numpy as np

# a = np.array([2, 3])
# b = np.array([4, 0])

# projection = (np.dot(a, b) / np.dot(b, b)) * b
# proj_length = abs(np.dot(a, b)) / np.linalg.norm(b)

# print("投影向量:", projection)
# print("投影长度:", proj_length)

# import numpy as np

# def project_point_on_segment(point, start, end):
# v = end - start
# w = point - start
# mu = np.dot(w, v) / np.dot(v, v)
# mu = np.clip(mu, 0, 1) # 限制在段内
# return start + mu * v

# p = np.array([4, 3])
# s = np.array([2, 2])
# e = np.array([6, 6])

# print("投影点:", project_point_on_segment(p, s, e))


# 平面 点法式
# Ax + By + Cz - (Ax0 + By0 + Cz0) = 0

# 给定的平面Plane的方程为:
# Ax+By+Cz+D=0
# x = xi - A*t
# y = yi - B*t
# z = zi - C*t
 
# t = (A*xi + B*yi + C*zi + D)/(A**2 + B**2 + C**2)

