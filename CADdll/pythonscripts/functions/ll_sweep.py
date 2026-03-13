from Autodesk.AutoCAD.DatabaseServices import SubentityId, SubentityType, FullSubentityPath, MeshFaceterData, IdMapping, ObjectIdCollection, SubDMesh, Curve, Extents3d, Polyline, Polyline3d, Line, Circle, Poly3dType, DBText, Region, BooleanOperationType
from Autodesk.AutoCAD.BoundaryRepresentation import PointContainment, Brep, Face, BrepEntity

from Autodesk.AutoCAD.Internal import Utils

import System


import clr

import acad
import academit

import System


def 命令(): 
    academit.添加命令("llsweep-pl-nj-to-pl-sheet-for", llsweep_pl_nj_to_pl_sheet_for)
    academit.添加命令("llsweep-pl-nj-to-pl-zj-rect-for", llsweep_pl_nj_to_pl_zj_rect_for)
    academit.添加命令("llsweep-subtract-ness-cube-for", llsweep_subtract_ness_cube_for)
    academit.添加命令("llsweep-rotation-to-flatten", llsweep_rotation_to_flatten)
    academit.添加命令("llsweep-flatten-to-rotation", llsweep_flatten_to_rotation)
    pass








@acad.decorator_command
def llsweep_pl_nj_to_pl_sheet_for(): 
    import ll_pl
    ll_pl.zhu_uipl_ness()
    pt1, pt2 = acad.GetPoint2("请选择基点: ", "请选择终点: ")
    if pt1 == None: return

    dr0 = acad.Direct(pt1, pt2)
    objidlist = acad.SSGetIdList([[0, "LWPOLYLINE"]])
    buflist = []
    for objid in objidlist:
        pt0 = acad.GetStartPoint(objid)
        pt0 = acad.Vec3Add(pt0, dr0)
        directlist = acad.GetLWPolyLineDirectList(objid)
        lengthlist = acad.GetLWPolyLineLengthList(objid)
        lengthlist = ll_pl.zhu_llpl_nj_to_zj(lengthlist)
        ptlist = ll_pl.zhu_llpl_build_pl(pt0, directlist, lengthlist)
        buflist.append(ptlist)

    with acad.transaction() as trans:
        for ptlist in buflist:
            pline1 = acad.AddLWPolyLine(ptlist) 
            dbobjrefcollect1 = pline1.GetOffsetCurves( ll_pl.zhu_llpl_ness_half)  
            dbobjrefcollect2 = pline1.GetOffsetCurves(-ll_pl.zhu_llpl_ness_half)  
            for objref1 in dbobjrefcollect1: acad.AddDBObject(objref1)
            for objref2 in dbobjrefcollect2: acad.AddDBObject(objref2)
            pline1.Erase()



@acad.decorator_command
def llsweep_pl_nj_to_pl_zj_rect_for(): 
    import ll_pl
    ll_pl.zhu_uipl_ness()
    pt1, pt2 = acad.GetPoint2("请选择基点: ", "请选择终点: ")
    if pt1 == None: return

    dr0 = acad.Direct(pt1, pt2)
    objidlist = acad.SSGetIdList([[0, "LWPOLYLINE"]])
    buflist = []
    for objid in objidlist:
        pt0 = acad.GetStartPoint(objid)
        pt0 = acad.Vec3Add(pt0, dr0)
        directlist = acad.GetLWPolyLineDirectList(objid)
        resultlist = ll_pl.zhu_llpl_nj_to_zj_sweep_rect(pt0, directlist)
        for ptlist in resultlist: 
            buflist.append(ptlist)

    with acad.transaction() as trans:
        for ptlist in buflist:
            acad.AddLWPolyLine(ptlist) 



@acad.decorator_command
def llsweep_subtract_ness_cube_for(): 
    import ll_pl
    ll_pl.zhu_uipl_ness()
    limit1 = ll_pl.zhu_llpl_ness + 0.05
    limit2 = ll_pl.zhu_llpl_ness_double + 0.05
    objidlist = acad.SSGetIdList([[0, "3DSOLID"]])
    with acad.transaction() as trans:
        reflist = []
        for objid in objidlist:
            objref = acad.TransObjectForWrite(objid)
            brep = Brep(objref)
            vertexlist = []
            for vertex in brep.Vertices:
                vertexlist.append([vertex.Point.X, vertex.Point.Y, vertex.Point.Z])

            count = len(vertexlist)
            indexlist = []
            resultlist = []
            for i in range(count):
                if i in indexlist: continue
                for j in range(i+1, count):
                    pt1 = vertexlist[i]
                    pt2 = vertexlist[j]
                    distance = acad.Distance(pt1, pt2)
                    if distance > limit1 and distance < limit2:
                        resultlist.append([pt1, pt2])
                        indexlist.append(i)
                        indexlist.append(j)

            if indexlist == []: continue
            [pt1, pt3], [po1, po3] = resultlist[0:2]
            mid1 = acad.MidPt1Pt2(pt1, pt3) 
            mid2 = acad.MidPt1Pt2(po1, po3) 
            
            buflist = []
            for face in brep.Faces:
                # 判断两点在面内
                vertexlist = []
                for loop in face.Loops:
                    for vertex in loop.Vertices: vertexlist.append([vertex.Point.X, vertex.Point.Y, vertex.Point.Z])
                if acad.IsPointIsPointListPoint(pt1, vertexlist) and acad.IsPointIsPointListPoint(pt3, vertexlist):
                    # 找出点和对向边
                    for loop in face.Loops:
                        for edge in loop.Edges:
                            pe1 = [edge.Vertex1.Point.X, edge.Vertex1.Point.Y, edge.Vertex1.Point.Z]
                            pe2 = [edge.Vertex2.Point.X, edge.Vertex2.Point.Y, edge.Vertex2.Point.Z]
                            if acad.IsPointSame(pt1, pe1) and acad.IsPointSame(pt3, pe2): continue
                            if acad.IsPointSame(pt3, pe1) and acad.IsPointSame(pt1, pe2): continue
                            if acad.IsPointSame(pt1, pe1) or acad.IsPointSame(pt1, pe2): buflist.append([pt3, pe1, pe2])
                            if acad.IsPointSame(pt3, pe1) or acad.IsPointSame(pt3, pe2): buflist.append([pt1, pe1, pe2])
            
            findlist = []
            for pt0, pe1, pe2 in buflist:
                line1 = acad.DBObjectLine(pe1, pe2)
                point = line1.GetClosestPointTo(acad.ToPoint3d(pt0), extend=False) # 会返回直线端点 # System.Boolean.Parse("False") 
                pi1 = [point.X, point.Y, point.Z]
                if acad.IsPointSame(pi1, pe1) or acad.IsPointSame(pi1, pe2): continue
                findlist.append(pi1)
            [pt2, pt4] = findlist
            # dr1 = acad.Direct(mid1, mid2)
            # [pt2, pt4] = acad.MatrixRotationPointList(90, dr1, mid1, [pt1, pt3])
            lineref1 = acad.AddLine(mid1, mid2)
            triref1 = acad.AddPolyline3d([pt1, pt3, pt2, pt1])
            triref2 = acad.AddPolyline3d([pt1, pt3, pt4, pt1])
            reflist.append([objid, lineref1, triref1, triref2])

    # large_objref.BooleanOperation(BooleanOperationType.BoolSubtract, objref)
    for objid0, lineref1, triref1, triref2 in reflist:
        objid1 = Utils.EntLast()
        ss1 = acad.SSSetFromIdList([triref1.ObjectId, triref2.ObjectId])   
        acad.Command(["SWEEP", ss1, "", lineref1.ObjectId]), acad.Prompt("\n") # 分开单独扫掠会出错
        objid2 = Utils.EntNext(objid1, skipSubEnt=True)
        objid3 = Utils.EntNext(objid2, skipSubEnt=True)
        # with acad.transaction() as trans:
        #     objref0 = acad.TransObjectForWrite(objid0)
        #     objref2 = acad.TransObjectForWrite(objid2)
        #     objref3 = acad.TransObjectForWrite(objid3)
        #     objref0.BooleanOperation(BooleanOperationType.BoolSubtract, objref2)
        #     objref0.BooleanOperation(BooleanOperationType.BoolSubtract, objref3)
        ss2 = acad.SSSetFromIdList([objid2, objid3])     
        # acad.SSSetFirst(ss2)
        acad.Command(["SUBTRACT", objid0, "", ss2, ""]), acad.Prompt("\n")



@acad.decorator_command
def llsweep_rotation_to_flatten(): 
    objidlist = acad.SSGetIdList([[0, "3DSOLID"]])
    while True:
        pt1, pt2, pt3 = acad.GetPoint3()
        if pt1 == None: break
        bufid = None
        for objid in objidlist:
            if acad.IsPointIsSolidVertex(pt1, objid):
                bufid = objid
                break
        if acad.IsNoneObjectId(bufid): break
        objidlist.remove(bufid)
        dr1 = acad.Direct(pt2, pt1)
        dr2 = acad.Direct(pt2, pt3)
        dr1 = acad.Vec3ResetLength(dr1, 10)
        dr2 = acad.Vec3ResetLength(dr2, 10)
        normal = acad.Cross(dr1, dr2)
        angle = acad.AngleFromDotDr1Dr2(dr1, dr2)
        po1 = pt2
        po2 = acad.Vec3Add(po1, normal)
        ss1 = acad.SSSetFromIdList(objidlist)
        acad.CommandRotate3d(ss1, po1, po2, 180-angle)

zhu_sweep_angle = 90
@acad.decorator_command
def llsweep_flatten_to_rotation(): 
    global zhu_sweep_angle 
    objidlist = acad.SSGetIdList([[0, "3DSOLID"]])
    while True:
        pt1, pt2, pt3 = acad.GetPoint3()
        if pt1 == None: break
        bufid = None
        for objid in objidlist:
            if acad.IsPointIsSolidVertex(pt1, objid):
                bufid = objid
                break
        if acad.IsNoneObjectId(bufid): break
        objidlist.remove(bufid)
        dr1 = acad.Direct(pt2, pt1)
        dr2 = acad.Direct(pt2, pt3)
        dr1 = acad.Vec3ResetLength(dr1, 10)
        dr2 = acad.Vec3ResetLength(dr2, 10)
        normal = acad.Cross(dr1, dr2)
        angle = acad.GetDouble(zhu_sweep_angle, "请输入折弯角度: ")
        if angle < 0  : angle = 0
        if angle > 180: angle = 180
        if angle != None: zhu_sweep_angle = angle
        po1 = pt2
        po2 = acad.Vec3Add(po1, normal)
        ss1 = acad.SSSetFromIdList(objidlist)
        acad.CommandRotate3d(ss1, po1, po2, 180-zhu_sweep_angle)


@acad.decorator_command
def llpl_sweep():
    plineid = acad.EntSel("请点击扫掠对象: ")
    rectid = acad.EntSel("请点击路径对象: ")
    xydrlist = acad.GetLWPolyLineDirectList(plineid)
    rectptlist = acad.GetLWPolyLinePointList(rectid)
    with acad.transaction() as trans:
        pt0 = acad.TransStartPoint(plineid)
        mid = acad.TransLWPolyLineStartMid(plineid)
        acad.AddText(pt0, "起点")
        acad.AddText(mid, "中点")
        pt0 = acad.TransStartPoint(rectid)
        mid = acad.TransLWPolyLineStartMid(rectid)
        acad.AddText(pt0, "起点")
        acad.AddText(mid, "中点") 

    # 偏移 + 垂直移动
    with acad.command_undo(), acad.command_osmode():
        dx, dz = 0, 0
        rectidlist = [rectid]
        for [x, y, z] in xydrlist:
            dx += x
            dz += y
            acad.CommandOffSet(rectid, acad.Absolute(dx), acad.Vec3Add(pt0, [dx,0,0]))
            lastid = acad.EntLast()
            rectidlist.append(lastid)
            acad.CommandMove(lastid, [0,0,0], [0,0,dz])

    result_list = []

    # 添加角点多段线       
    ptlist_list = []
    for rectid in rectidlist:
        pline_point_list = acad.GetLWPolyLinePointList(rectid)
        ptlist_list.append(pline_point_list)

    if ptlist_list == []: return
    count = len(ptlist_list[0]) 
    for i in range(count):
        ptlist = []
        for valuelist in ptlist_list:
            ptlist.append(valuelist[i])
        result_list.append(ptlist)


    # 添加直线端点多段线
    xzdrlist = []
    for [x, y, z] in xydrlist:
        xzdrlist.append([x, z, y])    

    for i in range(len(rectptlist)-1):
        pt1 = rectptlist[i]
        pt2 = rectptlist[i+1]
        match i:
            case 0: drlist = xzdrlist
            case 1: drlist = acad.ChangeCoordinateXY(xzdrlist, "-Y",  "X")
            case 2: drlist = acad.ChangeCoordinateXY(xzdrlist, "-X", "-Y")
            case 3: drlist = acad.ChangeCoordinateXY(xzdrlist,  "Y", "-X")
        pt1list = acad.DirectListToPointList(pt1, drlist)
        pt2list = acad.DirectListToPointList(pt2, drlist)
        result_list.append(pt1list)
        result_list.append(pt2list)


    # # 添加中点多段线    
    # ptlist_list = []
    # for rectid in rectidlist:
    #     pline_point_list = acad.GetLWPolyLineMidPointList(rectid)
    #     ptlist_list.append(pline_point_list)

    # if ptlist_list == []: return
    # count = len(ptlist_list[0]) 
    # for i in range(count):
    #     ptlist = []
    #     for valuelist in ptlist_list:
    #         ptlist.append(valuelist[i])
    #     result_list.append(ptlist)

    if result_list == []: return
    with acad.transaction() as trans:
        for ptlist in result_list:
            acad.AddPolyline3d(ptlist)



@acad.decorator_command
def llpl_sweep_set():
    plineid = acad.EntSel("请点击扫掠对象: ")
    rectid = acad.EntSel("请点击路径对象: ")
    xydrlist = acad.GetLWPolyLineDirectList(plineid)
    rectptlist = acad.GetLWPolyLinePointList(rectid)
    with acad.transaction() as trans:
        pt0 = acad.TransStartPoint(plineid)
        mid = acad.TransLWPolyLineStartMid(plineid)
        acad.AddText(pt0, "起点")
        acad.AddText(mid, "中点")
        pt0 = acad.TransStartPoint(rectid)
        mid = acad.TransLWPolyLineStartMid(rectid)
        acad.AddText(pt0, "起点")
        acad.AddText(mid, "中点") 

    # 偏移 + 垂直移动
    with acad.command_undo(), acad.command_osmode():
        dx, dz = 0, 0
        rectidlist = [rectid]
        for [x, y, z] in xydrlist:
            dx += x
            dz += y
            acad.CommandOffSet(rectid, acad.Absolute(dx), acad.Vec3Add(pt0, [dx,0,0]))
            lastid = acad.EntLast()
            rectidlist.append(lastid)
            acad.CommandMove(lastid, [0,0,0], [0,0,dz])

    result_list = []

    # 添加角点多段线       
    ptlist_list = []
    for rectid in rectidlist:
        pline_point_list = acad.GetLWPolyLinePointList(rectid)
        ptlist_list.append(pline_point_list)

    if ptlist_list == []: return
    count = len(ptlist_list[0]) 
    for i in range(count):
        ptlist = []
        for valuelist in ptlist_list:
            ptlist.append(valuelist[i])
        result_list.append(ptlist)


    # 添加直线端点多段线
    xzdrlist = []
    for [x, y, z] in xydrlist:
        xzdrlist.append([x, z, y])    

    for i in range(len(rectptlist)-1):
        pt1 = rectptlist[i]
        pt2 = rectptlist[i+1]
        match i:
            case 0: drlist = xzdrlist
            case 1: drlist = acad.ChangeCoordinateXY(xzdrlist, "-Y",  "X")
            case 2: drlist = acad.ChangeCoordinateXY(xzdrlist, "-X", "-Y")
            # case 3: drlist = acad.ChangeCoordinateXY(xzdrlist,  "Y", "-X")
            case _: break
        pt1list = acad.DirectListToPointList(pt1, drlist)
        pt2list = acad.DirectListToPointList(pt2, drlist)
        result_list.append(pt1list)
        result_list.append(pt2list)


    # # 添加中点多段线    
    # ptlist_list = []
    # for rectid in rectidlist:
    #     pline_point_list = acad.GetLWPolyLineMidPointList(rectid)
    #     ptlist_list.append(pline_point_list)

    # if ptlist_list == []: return
    # count = len(ptlist_list[0]) 
    # for i in range(count):
    #     ptlist = []
    #     for valuelist in ptlist_list:
    #         ptlist.append(valuelist[i])
    #     result_list.append(ptlist)

    if result_list == []: return
    with acad.transaction() as trans:
        for ptlist in result_list:
            acad.AddPolyline3d(ptlist)


# @acad.decorator_command
# def llpl_loft():
#     pass
#     # objidlist = acad.SSGetIdList([[0, "LWPOLYLINE"]])
#     # with acad.transaction() as trans:
#     #     for objid in objidlist:
#     #         acad.Copy(objid, [0,0], [1000,1000])


# // We use a builder object to create
# // our SweepOptions
# SweepOptionsBuilder sob = new SweepOptionsBuilder();
# // Align the entity to sweep to the path
# sob.Align = SweepOptionsAlignOption.AlignSweepEntityToPath;
# // The base point is the start of the path
# sob.BasePoint = pathEnt.StartPoint;
# // The profile will rotate to follow the path
# sob.Bank = true;
# // Now generate the solid or surface...
# Entity ent;
# if (createSolid)
# {
# Solid3d sol = new Solid3d();
# sol.CreateSweptSolid( sweepEnt, pathEnt, sob.ToSweepOptions());
# ent = sol;
# }
# else
# {
# SweptSurface ss = new SweptSurface();
# ss.CreateSweptSurface( sweepEnt, pathEnt, sob.ToSweepOptions() );
# ent = ss;
# }
