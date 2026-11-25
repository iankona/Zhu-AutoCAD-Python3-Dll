from Autodesk.AutoCAD.DatabaseServices import SubentityId, SubentityType, FullSubentityPath, MeshFaceterData, IdMapping, ObjectIdCollection, SubDMesh, Curve, Extents3d, Polyline, Polyline3d, Line, Circle, Poly3dType, DBText, Region, BooleanOperationType
from Autodesk.AutoCAD.BoundaryRepresentation import PointContainment, Brep, Face, BrepEntity
import System



import re
import acad
import academit

def 命令(): 
    academit.添加命令("ll-solid-up", ll_solid_up)
    academit.添加命令("wall", wall)
    academit.添加命令("wallrec", wallrec)
    academit.添加命令("wallrecmove", wallrecmove)
    # academit.添加命令("wallzup", wallrecmove)
    academit.添加命令("wallpl2", wallpl2)
    academit.添加命令("holerec", holerec)
    academit.添加命令("wind", wind)
    academit.添加命令("windrec", windrec)
    academit.添加命令("recregion", recregion)
    academit.添加命令("yayin", yayin)
    academit.添加命令("planefromzpoint", facefromzpoint)

    academit.添加命令("wall-connet-all-to-zup", wall_connet_all_to_zup)


    academit.添加命令("ll-solid-subcut", ll_solid_subcut)
    academit.添加命令("ll-solid-intersect", ll_solid_intersect)

    academit.添加命令("ll-solid-subentity", ll_solid_subentity)
    academit.添加命令("ll-solid-slice3", ll_solid_slice3)
    academit.添加命令("ll-solid-slice2", ll_solid_slice2)







@acad.decorator_command
def ll_solid_slice2():
    while True:
        objidlist = acad.SSGetIdList([[0, "3DSOLID"]], string="请选择2条相交的实体:")
        if objidlist == []: return
        with acad.transaction() as trans:
            solid1, solid2 = acad.TransObjectIdListForWrite(objidlist[0:2])
            solid3, solid4 = solid1.Clone(), solid2.Clone()
            solid3.BooleanOperation(BooleanOperationType.BoolIntersect, solid4)
            acad.AddDBObject(solid3)
        pt1, pt2 = acad.GetPoint2(baseptflag=True)
        if pt1 == None: return
        with acad.transaction() as trans:
            solid1, solid2, solid3 = acad.TransObjectIdListForWrite([solid1.ObjectId, solid2.ObjectId, solid3.ObjectId])

            pt3 = solid_sheet_pt1pt2_find_pt3(solid1, pt1, pt2)
            plane1 = acad.Plane(acad.ToPoint3d(pt1), acad.ToPoint3d(pt2), acad.ToPoint3d(pt3))
            solid1 = solid1.Slice(plane1, True)

            pt3 = solid_sheet_pt1pt2_find_pt3(solid2, pt1, pt2)
            plane2 = acad.Plane(acad.ToPoint3d(pt1), acad.ToPoint3d(pt2), acad.ToPoint3d(pt3))
            solid2 = solid2.Slice(plane2, True)
            
            acad.AddDBObject(solid1)
            acad.AddDBObject(solid2)
            solid3.Erase()


def solid_sheet_pt1pt2_find_pt3(solid, pt1, pt2):
    brep = acad.BrepWithSubEntityFromRef(solid)
    dislist = []
    for edge in brep.Edges:
        curve = acad.BrepEdgeToDBObjectCurve(edge)
        point = curve.GetClosestPointTo(acad.ToPoint3d(pt2), extend=False)
        po1 = [point.X, point.Y, point.Z]
        dis = acad.Distance(pt2, po1)
        dislist.append([po1, dis])

    dislist.sort(key = lambda item: item[1], reverse=False) # 排序规则，reverse = True 降序， reverse = False 升序（默认）
    # dr1 = acad.Direct(pt1, pt2)
    # dr2 = acad.Direct(pt2, po1)
    # angle = acad.AngleFromCrossDr1Dr2(dr1, dr2)
    # if 89 < angle < 91: return po1
    # print(dislist)
    return dislist[1][0]
    







@acad.decorator_command
def ll_solid_slice3():
    while True:
        objidlist = acad.SSGetIdList([[0, "3DSOLID"]], string="请选择2条相交的实体:")
        if objidlist == []: return
        with acad.transaction() as trans:
            solid1, solid2 = acad.TransObjectIdListForWrite(objidlist[0:2])
            solid3, solid4 = solid1.Clone(), solid2.Clone()
            solid3.BooleanOperation(BooleanOperationType.BoolIntersect, solid4)
            acad.AddDBObject(solid3)
        pt1, pt2, pt3 = acad.GetPoint3(baseptflag=True)
        if pt1 == None: return
        with acad.transaction() as trans:
            solid1, solid2, solid3 = acad.TransObjectIdListForWrite([solid1.ObjectId, solid2.ObjectId, solid3.ObjectId])
            plane = acad.Plane(acad.ToPoint3d(pt1), acad.ToPoint3d(pt2), acad.ToPoint3d(pt3))
            solid1 = solid1.Slice(plane, True)
            solid2 = solid2.Slice(plane, True)
            acad.AddDBObject(solid1)
            acad.AddDBObject(solid2)
            solid3.Erase()




@acad.decorator_command
def ll_solid_subentity():
    objid = acad.EntSel([[0, "3DSOLID"]])
    brep = acad.BrepWithSubEntityFromId(objid)
    result = []
    for face in brep.Faces:
        for  edge in face.Brep.Edges:
            result.append(edge.SubentityPath.SubentId)
    # print(result) # It Works!!!


# @acad.decorator_command
# def ll_solid_subcut():
#         objid = acad.EntSel([[0, "3DSOLID"]], string="请选择2条相交的实体:")
#         if acad.IsNone(objid): return
#         pt1, 
#         with acad.transaction() as trans:
#             solid1, solid2 = acad.TransObjectIdListForWrite(objidlist[0:2])
#             solid1.BooleanOperation(BooleanOperationType.BoolSubtract, solid2.Clone())
#             solid2.BooleanOperation(BooleanOperationType.BoolSubtract, solid1.Clone())


# public enum BooleanOperationType {
#   BoolUnite,
#   BoolIntersect,
#   BoolSubtract
# }

@acad.decorator_command
def ll_solid_subcut():
    while True:
        objidlist = acad.SSGetIdList([[0, "3DSOLID"]], string="请选择2条相交的实体:")
        if objidlist == []: return
        with acad.transaction() as trans:
            solid1, solid2 = acad.TransObjectIdListForWrite(objidlist[0:2])
            solid1.BooleanOperation(BooleanOperationType.BoolSubtract, solid2.Clone())
            solid2.BooleanOperation(BooleanOperationType.BoolSubtract, solid1.Clone())



@acad.decorator_command
def ll_solid_intersect():
    while True:
        objidlist = acad.SSGetIdList([[0, "3DSOLID"]], string="请选择2条相交的实体:")
        if objidlist == []: return
        with acad.transaction() as trans:
            solid1, solid2 = acad.TransObjectIdListForWrite(objidlist[0:2])
            solid1, solid2 = solid1.Clone(), solid2.Clone()
            solid1.BooleanOperation(BooleanOperationType.BoolIntersect, solid2)
            acad.AddDBObject(solid1)




@acad.decorator_command
def wallpl2():
    ui_wall_height()
    while True:
        objidlist = acad.SSGetIdList([[0, "LWPOLYLINE"]], string="请选择2条多段线:")
        pl1list = acad.GetMKPolyLinePointList(objidlist[0])
        pl2list = acad.GetMKPolyLinePointList(objidlist[1])
        pt1, pt2 = pl1list[0], pl1list[-1]
        po1, po2 = pl2list[0], pl2list[-1]
        if acad.Distance(pt1, po1) < acad.Distance(pt1, po2):
            ptlist = pl1list + pl2list[::-1]
        else:
            ptlist = pl1list + pl2list
        if ptlist == []: return
        extrude_from_region(ptlist, zhu_extrude_height) # ptlist_to_region(ptlist)   




def objid_find_connet_objid1_objid2():
    pass






@acad.decorator_command
def wall_connet_all_to_zup():
    objidlist = acad.SSGetIdList([[0, "REGION"]], string="请选择内墙面:")
    objid = acad.EntSel([[0, "REGION"]], string="请选择基准面:")
    pt1 = acad.GetPoint("请选择基准点:")
    with acad.transaction() as trans:
        objreflist = acad.TransObjectIdListForRead(objidlist)



    
    pc0 = acad.GetIdListBoundCenterXYZ(objidlist)
    objid_dr_list = []
    for objid, objref in zip(objidlist, objreflist):
        pn0 = objref.Normal
        pn1 = acad.PointToPlaneNPoint(pc0)
        drn = acad.Direct(pn1, pc0)
        objid_dr_list.append([objid, drn])

    
        
    





@acad.decorator_command
def facefromzpoint():
    pt1, pt2, pt3 = acad.GetPoint3("请选择平面第1点：", "请选择平面第2点：", "请选择平面第3点：", baseptflag=True)
    if pt1 == None: return
    ptlist = acad.GetPointX(baseptflag=True)
    if ptlist == []: return 
    pn0, normal = acad.PlaneFromPt1Pt2Pt3(pt1, pt2, pt3)
    buflist = []
    for po1 in ptlist:
        x1, y1, z1 = po1
        pt1 = [x1, y1, z1]
        pt2 = [x1, y1, z1+1000]
        pd1 = acad.LineToPlanePoint(pt1, pt2, pn0, normal)
        print(po1, pd1)
        buflist.append(pd1)
    if buflist == []: return
    ptlist_to_region(buflist)

    






@acad.decorator_command
def yayin():
    while True:
        pt1, pt2 = acad.GetPoint2(baseptflag=True)
        if pt1 == None: break
        objid = acad.GetSelectPickId(pt1)
        with acad.transaction() as trans:
            line1 = acad.AddLine(pt1, pt2)
        # ss1 = acad.SSSetFromIdList([objid])
        acad.Command(["imprint", objid, line1.ObjectId, "Y", ""])




zhu_extrude_height = 2600
def ui_wall_height():
    global zhu_extrude_height 
    height = acad.GetDouble(zhu_extrude_height, "请输入墙高:")
    if height != None: zhu_extrude_height = height


zhu_move_height = 2600
def ui_wall_move_height():
    global zhu_extrude_height, zhu_move_height
    height1 = acad.GetDouble(3000, "请输入第1个高度:")
    height2 = acad.GetDouble(2600, "请输入第2个高度:")
    height = abs(height1 - height2)
    zhu_extrude_height = height
    zhu_move_height = height2
    # height = acad.GetDouble(zhu_move_height, "请输入移动高度:")
    # if height != None: zhu_move_height = height


zhu_hole_height = 2000
def ui_hole_height():
    global zhu_hole_height 
    height = acad.GetDouble(zhu_hole_height, "请输入门梁高:")
    if height != None: zhu_hole_height = height



def ptlist_to_region(ptlist):
    pt1, pt2, pt3 = ptlist[0:3]
    dr1 = acad.Direct(pt1, pt2)
    dr2 = acad.Direct(pt2, pt3)
    nor = acad.Cross(dr1, dr2)
    if nor[2] < 0: ptlist = ptlist[::-1]
    with acad.transaction() as trans:
        buflist = []
        count = len(ptlist)
        for i in range(count):
            k = i % count
            m = (i+1) % count
            pt1 = ptlist[k]
            pt2 = ptlist[m]
            line1 = acad.DBObjectLine(pt1, pt2)
            buflist.append(line1)

        collect = acad.DBObjectCollection()
        for line1 in buflist: collect.Add(line1)
        region = acad.Region.CreateFromCurves(collect)[0]
        acad.AddDBObject(region)
    return region

def extrude_from_region(ptlist, height):
    pt1, pt2, pt3 = ptlist[0:3]
    dr1 = acad.Direct(pt1, pt2)
    dr2 = acad.Direct(pt2, pt3)
    nor = acad.Cross(dr1, dr2)
    if nor[2] < 0: ptlist = ptlist[::-1]
    with acad.transaction() as trans:
        buflist = []
        count = len(ptlist)
        for i in range(count):
            k = i % count
            m = (i+1) % count
            pt1 = ptlist[k]
            pt2 = ptlist[m]
            line1 = acad.DBObjectLine(pt1, pt2)
            buflist.append(line1)

        collect = acad.DBObjectCollection()
        for line1 in buflist: collect.Add(line1)
        region = acad.Region.CreateFromCurves(collect)[0]
        
        solid = None
        if height > 0:
            solid = acad.Solid3d()
            solid.Extrude(region, height, 0)
            acad.AddDBObject(solid)
    return solid



def get_ptlist():
    ptlist = []
    while True:
        pt1 = acad.GetPoint()
        if pt1 == None: break 
        ptlist.append(pt1)
    return ptlist


def get_ptlist_rec():
    ptlist = []
    pt1, pt3 = acad.GetCorner2("请选择对角点1:", "请选择对角点2:")
    if pt1 != None: ptlist = acad.Pt1Pt3toPt1Pt2Pt3Pt4XY0(pt1, pt3)
    return ptlist


@acad.decorator_command
def wall():
    ui_wall_height()
    while True:
        ptlist = get_ptlist()
        if ptlist == []: return
        extrude_from_region(ptlist, zhu_extrude_height)




@acad.decorator_command
def recregion():
    while True:
        pt1, pt3 = acad.GetCorner2("请选择对角点1:", "请选择对角点2:")
        if pt1 == None: return
        ptlist = acad.Pt1Pt3toPt1Pt2Pt3Pt4XYZ(pt1, pt3)
        if ptlist == []: return
        ptlist_to_region(ptlist)   



@acad.decorator_command
def wallrec():
    ui_wall_height()
    while True:
        pt1, pt3 = acad.GetCorner2("请选择对角点1:", "请选择对角点2:")
        if pt1 == None: return
        ptlist = acad.Pt1Pt3toPt1Pt2Pt3Pt4XY0(pt1, pt3)
        if ptlist == []: return
        extrude_from_region(ptlist, zhu_extrude_height)            
            

@acad.decorator_command
def wallrecmove():
    ui_wall_move_height()
    while True:
        pt1, pt3 = acad.GetCorner2("请选择对角点1:", "请选择对角点2:")
        if pt1 == None: return
        ptlist = acad.Pt1Pt3toPt1Pt2Pt3Pt4XY0(pt1, pt3)
        if ptlist == []: return
        solid = extrude_from_region(ptlist, zhu_extrude_height) 
        with acad.transaction() as trans:
            acad.TransMove(solid.ObjectId, sourcept=[0,0,0], targetpt=[0,0,zhu_move_height])


@acad.decorator_command
def holerec():
    while True:
        objid = acad.EntSel([[0, "3DSOLID"]], "请选择墙体:")
        ui_hole_height()
        pt1, pt3 = acad.GetCorner2("请选择对角点1:", "请选择对角点2:")
        if pt1 == None: break 
        ptlist = acad.Pt1Pt3toPt1Pt2Pt3Pt4XY0(pt1, pt3)
        with acad.transaction():
            objref = acad.TransObjectForWrite(objid)
            solid = extrude_from_region(ptlist, zhu_hole_height)
            objref.BooleanOperation(BooleanOperationType.BoolSubtract, solid)


def get_text_string():
    objid = acad.EntSel([[-4, "<OR"],[0, "TEXT"],[0, "MTEXT"], [-4, "OR>"]], "请选择文本:")
    string = None
    with acad.transaction():
        objref = acad.TransObjectForRead(objid)
        match str(objref): 
            case "Autodesk.AutoCAD.DatabaseServices.DBText": string = objref.TextString
            case "Autodesk.AutoCAD.DatabaseServices.MText": string = objref.Text # 只读对像,无格式 # objref.Contents是文本写入对像 # 'SimSun|b0|i0|c134|p2;窗下:}525mm\\P{\\fSimSun|b0|i0|c134|p2;窗高:1235m'
    return string


def text_to_height(string):
    # string = string.replace("\r", "")
    # lines = string.split("\n")
    # a = None
    # b = None
    # for line in lines:
    #     if "窗下" in line: a = line
    #     if "窗高" in line: b = line
    # if a == None or b == None: return None
    # a = a.replace(" ", "")
    # b = b.replace(" ", "")
    # print(a.encode('utf-8'), b.encode('utf-8'))
    # # b'\xe7\xaa\x97\xe4\xb8\x8b:525mm\r' b'\xe7\xaa\x97\xe9\xab\x98:1235mm'
    # a = a[3:-2]
    # b = b[3:-2]
    result = re.findall(r'\d+', string)
    # print(string, result)
    if len(result) == 1:
        return 0, float(result[0])
    return float(result[1]), float(result[0])





@acad.decorator_command
def wind():
    ui_wall_height()
    text = get_text_string()
    if text == None: return
    height1, height2 = text_to_height(text)
    height3 = zhu_extrude_height - height1 - height2
    ptlist = get_ptlist()
    if ptlist == None: return
    solid1 = extrude_from_region(ptlist, height1)
    # solid2 = extrude_from_region(ptlist, height2)
    solid3 = extrude_from_region(ptlist, height3)

    with acad.transaction():
        # solid2 = acad.TransObjectForWrite(solid2.ObjectId)
        solid3 = acad.TransObjectForWrite(solid3.ObjectId)

        # vecdr = acad.Vector3d(0,0, height1)
        # matrix4x4 = acad.Matrix3d.Displacement(vecdr)
        # solid2.TransformBy(matrix4x4)

        vecdr = acad.Vector3d(0,0, height1+height2)
        matrix4x4 = acad.Matrix3d.Displacement(vecdr)
        solid3.TransformBy(matrix4x4)
    


@acad.decorator_command
def windrec():
    ui_wall_height()
    text = get_text_string()
    if text == None: return
    height1, height2 = text_to_height(text)
    height3 = zhu_extrude_height - height1 - height2
    pt1, pt3 = acad.GetCorner2("请选择对角点1:", "请选择对角点2:")
    if pt1 == None: return
    ptlist = acad.Pt1Pt3toPt1Pt2Pt3Pt4XY0(pt1, pt3)
    solid1 = extrude_from_region(ptlist, height1)
    # solid2 = extrude_from_region(ptlist, height2)
    solid3 = extrude_from_region(ptlist, height3)

    with acad.transaction():
        # solid2 = acad.TransObjectForWrite(solid2.ObjectId)
        solid3 = acad.TransObjectForWrite(solid3.ObjectId)

        # vecdr = acad.Vector3d(0,0, height1)
        # matrix4x4 = acad.Matrix3d.Displacement(vecdr)
        # solid2.TransformBy(matrix4x4)

        vecdr = acad.Vector3d(0,0, height1+height2)
        matrix4x4 = acad.Matrix3d.Displacement(vecdr)
        solid3.TransformBy(matrix4x4)




def objidlist_pop_objidlist(objidlist, popidlist):
    buflist = [ str(objid) for objid in popidlist]
    cuflist = []
    for objid in objidlist:
        if str(objid) in buflist: continue
        cuflist.append(objid)
    return cuflist
    
def edgelistlist_pop_objidlist(edgelistlist, popidlist):
    buflist = [ str(objid) for objid in popidlist]
    cuflist = []
    for objid, edgelist in edgelistlist:
        if str(objid) in buflist: continue
        cuflist.append([objid, edgelist])
    return cuflist

def objid_to_edgelist(objid):
    objref = acad.GetObjectForRead(objid)
    brep = Brep(objref)
    edgelist = []
    for edge in brep.Edges:
        point1 = edge.Vertex1.Point
        point2 = edge.Vertex2.Point
        pt1 = [point1.X, point1.Y, point1.Z]
        pt2 = [point2.X, point2.Y, point2.Z]
        edgelist.append([pt1, pt2])
    return edgelist

def objidlist_to_edgelistlist(objidlist):
    result = []
    for objid in objidlist:
        objref = acad.GetObjectForRead(objid)
        brep = Brep(objref)
        edgelist = []
        for edge in brep.Edges:
            point1 = edge.Vertex1.Point
            point2 = edge.Vertex2.Point
            pt1 = [point1.X, point1.Y, point1.Z]
            pt2 = [point2.X, point2.Y, point2.Z]
            edgelist.append([pt1, pt2])
        result.append([objid, edgelist])
    return result


def find_connect_solid_zup(objid0, edgelistlist):
    # edgelist0 = None
    buflist = []
    for objid1, edgelist in edgelistlist:
        if str(objid0) == str(objid1): 
            edgelist0 = edgelist
            continue
        buflist.append([objid1, edgelist])


    cuflist = []
    duflist = []
    for objid1, edgelist1 in buflist:
        samelist = []
        flag = False
        for pt1, pt2 in edgelist0:
            for po1, po2 in edgelist1:
                if acad.IsEdgeSame(pt1, pt2, po1, po2): 
                    samelist.append([pt1, pt2])
                    flag = True
        if len(samelist) < 2: flag = False
        if flag: 
            cuflist.append([objid1, edgelist1])
        else:
            duflist.append([objid1, edgelist1])
    return cuflist, duflist



def is_face_and_face_connect(edgelist1, edgelist2):
    samelist = []
    for pt1, pt2 in edgelist1:
        for po1, po2 in edgelist2:
            if acad.IsEdgeSame(pt1, pt2, po1, po2): samelist.append([po1, po2])
    if len(samelist) >= 2: return True
    return False

def find_edgelist_and_edgelist_connect(edgelist1, edgelist2): # 默认已经判断过是相连的
    samelist = []
    for pt1, pt2 in edgelist1:
        for po1, po2 in edgelist2:
            if acad.IsEdgeSame(pt1, pt2, po1, po2): samelist.append([pt1, pt2])
    return samelist

def find_link_idlist(edgelist, edgelistlist):
    baselist = edgelist
    linklist = []
    indexlist = []
    for _ in range(len(edgelistlist)):
        hasfind = False
        for i, [objid1, edgelist] in enumerate(edgelistlist):
            if i in indexlist: continue
            if is_face_and_face_connect(baselist, edgelist): 
                hasfind = True
                baselist = edgelist
                indexlist.append(i)
                linklist.append(objid1)
        if not hasfind: break

    return linklist



def edgelist_to_solid_center(edgelist):
    count = len(edgelist) * 2
    x0, y0, z0 = 0, 0, 0
    for pt1, pt2 in edgelist:
        x1, y1, z1 = pt1
        x2, y2, z2 = pt2
        x0 += (x1+x2)
        y0 += (y1+y2)
        z0 += (z1+z2)
    return [x0/count, y0/count, z0/count]




def find_rotate_axis_and_angle(pt1, pt2, samelist, center):
    dr0 = acad.Direct(pt1, pt2) # 没有捕捉的条件下，选中物体时的点不在实体上，导致pt2，不是在预计的位置，有时向上，实际上却是比物体还低
    pt1 = center
    pt2 = acad.Vec3Add(center, dr0)

    # axis
    po1, po2 = samelist[0]
    po3, po4 = samelist[1]
    dis1 = acad.Distance(pt2, po1) + acad.Distance(pt2, po2)
    dis2 = acad.Distance(pt2, po3) + acad.Distance(pt2, po4)
    if dis1 > dis2:
        po1, po2, po3, po4 = po3, po4, po1, po2

    if acad.Distance(po1, po4) > acad.Distance(po1, po3):
        po3, po4 = po4, po3
    
    # ptlist = [po1, po2, po3, po4]
    # if not acad.IsCCWUseZup(ptlist): po1, po2, po3, po4 = po2, po1, po4, po3

    axis = acad.Direct(po1, po2)
    dr1 = acad.Direct(po1, po2)
    dr2 = acad.Direct(po2, po3)
    normal = acad.CrossNormalized(dr1, dr2)
    angle = acad.AngleFromDotDr1Dr2(normal, acad.Direct(pt1, pt2))
    if acad.AngleFromDotDr1Dr2(normal, acad.Direct(po1, center)) > 90: axis = acad.Direct(po2, po1)
    return po1, po2, axis, angle



def samelist_sort(samelist):
    euflist = []
    for pt1, pt2 in samelist:
        distance = acad.Distance(pt1, pt2)
        euflist.append([distance, pt1, pt2])
    euflist.sort(key = lambda item: item[0], reverse=True) # 从大到小

    samelist = []
    for i in range(2):
        a, pt1, pt2 = euflist[i]
        samelist.append([pt1, pt2])
    return samelist



def auto_rotation(pt1, pt2, objid0, objid1, linkidlist):
    edgelist0 = objid_to_edgelist(objid0)
    edgelist1 = objid_to_edgelist(objid1)
    samelist = find_edgelist_and_edgelist_connect(edgelist0, edgelist1)
    samelist = samelist_sort(samelist)
    center = edgelist_to_solid_center(edgelist1)
    po1, po2, axis, angle = find_rotate_axis_and_angle(pt1, pt2, samelist, center)
    with acad.transaction() as trans:
        bufidlist = [objid1] + linkidlist
        for objid2 in bufidlist:
            acad.TransRoation(objid2, angle, axis, po1)



@acad.decorator_command
def ll_solid_up():
    objidlist = acad.SSGetIdList()  
    edgelistlist = objidlist_to_edgelistlist(objidlist)
    while True:
        pt1, objid0 = acad.EntSelEntity()
        pt2 = acad.GetPoint(base_point=pt1)
        if pt2 == None: return
        connectlist, remainedgelistlist = find_connect_solid_zup(objid0, edgelistlist)  
        for objid1, edgelist in connectlist:
            linkidlist = find_link_idlist(edgelist, remainedgelistlist)
            auto_rotation(pt1, pt2, objid0, objid1, linkidlist)    
        popidlist = [objid0] # + [objid1 for objid1, samelist, edgelist in connectlist]
        objidlist = objidlist_pop_objidlist(objidlist, popidlist)
        edgelistlist = edgelistlist_pop_objidlist(edgelistlist, popidlist)
        # with acad.transaction() as trans:
        #     acad.AddLine(pt1, pt2)
        #     acad.AddText(pt2, "2", 5)




# @acad.decorator_command
# def ll_solid_up_v0():
#     objidlist = acad.SSGetIdList()  
#     edgelistlist = objidlist_to_edgelistlist(objidlist)
#     while True:
#         pt1, objid0 = acad.EntSelEntity()
#         pt2 = acad.GetPoint(base_point=pt1)
#         if pt2 == None: return
#         connectlist, remainedgelistlist = find_connect_solid_zup(objid0, edgelistlist)  
#         for objid1, edgelist in connectlist:
#             samelist = samelist_sort(samelist)
#             po1, po2, axis, normal = find_rotate_axis_and_normal(pt1, pt2, samelist)
#             angle = acad.AngleFromDotDr1Dr2(acad.Direct(pt1, pt2), normal)
#             linkidlist = find_link_idlist(edgelist, remainedgelistlist)
#             with acad.transaction() as trans:
#                 bufidlist = [objid1] + linkidlist
#                 for objid2 in bufidlist:
#                     acad.TransRoation(objid2, angle, axis, po1)
#         popidlist = [objid0] # + [objid1 for objid1, samelist, edgelist in connectlist]
#         objidlist = objidlist_pop_objidlist(objidlist, popidlist)


@acad.decorator_command
def ll_select_solid_edge():
    viewrecode = acad.ed.GetCurrentView() # ViewTableRecord
    print(viewrecode)
    objid = viewrecode.Layout
    print(objid)
    with acad.transaction() as trans:
        objref = acad.TransObjectForWrite(objid)
    acad.Prompt(objref.GetType())







@acad.decorator_command
def ll_brep_test():
    [pickpoint, objid] = acad.EntSelEntity()
    fullpath = FullSubentityPath([objid], SubentityId(SubentityType.Null, System.IntPtr.Zero))
    # objidlist = acad.SSGetIdList()
    # print(objidlist)
    # fullpath = FullSubentityPath(objidlist, SubentityId(SubentityType.Null, System.IntPtr.Zero)) # 只有第1个objid起作用
    brep = Brep(fullpath)
    for complex in brep.Complexes:
        print("complex")
        for shell in complex.Shells:
            print("shell")
            for face in shell.Faces:
                print("face")

# @acad.decorator_command
# def ll_imprint():
#     objid = acad.EntSel()
#     pt1, pt2 = acad.GetPoint2()
#     if pt1 == None: return 
#     with acad.transaction() as trans:
#         line = acad.AddLine(pt1, pt2)
#         objref = acad.TransObjectForWrite(objid)
#         brep = Brep(objref)
#         brep.Surf.ProjectOnToSurface(line, acad.Vector3d(0, 0, 1))


@acad.decorator_command
def ll_imprint():
    objid = acad.EntSel()
    while True:
        pt1, pt2 = acad.GetPoint2()
        if pt1 == None: return 
        with acad.transaction() as trans:
            line = acad.AddLine(pt1, pt2)
        acad.Command(["imprint", objid, line.ObjectId, "Y", ""])

# opts = new PromptSelectionOptions();
# opts.AllowSubSelections = true;
# AllowSubSelections allows to select sub entities (solid or meshes edges, attribute references, ...) accordingly with the LEGACYCTRLPICK and SUBOBJSELECTMODE sysvar settings.
# With default settings (LEGACYCTRLPICK = 2 and SUBOBJSELECTMODE = 0), AllowSubSelections allows sub entities selection using Ctrl+click.
# 你好
# AllowSubSelections 允许通过 LEGACYCTRLPICK 和 SUBOBJSELECTMODE sysvar 设置，相应地选择子实体（实体或网格、边缘、属性引用等）。
# 在默认设置（LEGACYCTRLPICK = 2 和 SUBOBJSELECTMODE = 0）下，允许子实体选择时，可以通过 Ctrl+click 选择子实体。





#     Reg1 = (Region)Regions[0];
# TmpSol.Extrude(Reg1, 9.798, 0.6154);

# Solid3d CylObj = new Solid3d();

# Region Hole = null;
# CircObj = new Circle(new Point3d(0, 0, 0), Reg1.Normal, 0.75);
# Curves.Clear();
# Curves.Add(CircObj);
# Regions.Clear();
# Regions = Region.CreateFromCurves(Curves);
# Hole = (Region)Regions[0];
# CylObj.Extrude(Hole, 10, 0);
# TmpSol.BooleanOperation(BooleanOperationType.BoolSubtract, CylObj);
# MdlSpc.AppendEntity(TmpSol);
# trans.AddNewlyCreatedDBObject(TmpSol, true);
# trans.Commit();