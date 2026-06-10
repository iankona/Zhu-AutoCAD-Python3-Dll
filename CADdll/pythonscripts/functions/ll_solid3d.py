from Autodesk.AutoCAD.DatabaseServices import SubentityId, SubentityType, FullSubentityPath, MeshFaceterData, IdMapping, ObjectIdCollection, SubDMesh, Curve, Extents3d, Polyline, Polyline3d, Line, Circle, Poly3dType, DBText, Region, BooleanOperationType
from Autodesk.AutoCAD.BoundaryRepresentation import PointContainment, Brep, Face, BrepEntity
import System

import acad
import academit

def 命令(): 
    academit.添加命令("ll-solid-up", ll_solid_up)


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