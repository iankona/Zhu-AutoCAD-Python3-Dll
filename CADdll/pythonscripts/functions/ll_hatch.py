from Autodesk.AutoCAD.DatabaseServices import SubentityId, SubentityType, FullSubentityPath, MeshFaceterData, IdMapping, ObjectIdCollection, SubDMesh, Curve, Extents3d, Polyline, Polyline3d, Line, Circle, Poly3dType, DBText, Region, BooleanOperationType
from Autodesk.AutoCAD.BoundaryRepresentation import PointContainment, Brep, Face, BrepEntity
import System

import acad
import academit

def 命令(): 
    academit.添加命令("ll-hatch-cirle", ll_hatch_circle)



def __point_in_region(pt1, brep):
    flag = False
    result = brep.GetPointContainment(acad.ToPoint3d(pt1), acad.PointContainment.Outside) # out ref 被自动处理成结果返回
    # print(brep_entity.GetType())
    # (<Autodesk.AutoCAD.BoundaryRepresentation.Face object at 0x00000204409347C0>, <PointContainment.OnBoundary: 2>)
    # (None, <PointContainment.Outside: 1>)
    if result[1] == acad.PointContainment.OnBoundary: 
        flag = True
    return flag


def __ptlist_in_region(ptlist, brep):
    flag = True
    for pt1 in ptlist:
        result = brep.GetPointContainment(acad.ToPoint3d(pt1), acad.PointContainment.Outside) # out ref 被自动处理成结果返回
        if result[1] != acad.PointContainment.OnBoundary: # (None, <PointContainment.Outside: 1>)
            flag = False
            break
    return flag


def __ptlist_add(ptlist, dr1):
    result = []
    for pt1 in ptlist:
        pt2 = acad.Vec3Add(pt1, dr1)
        result.append(pt2)
    return result

def __make_rectangle_mesh(pt1, pt2, distance):
    x1, y1, z1 = pt1
    x2, y2, z2 = pt2
    xlist = [x1]
    while True:
        x1 += distance
        if x1 < x2: 
            xlist.append(x1)
        else:
            break
    ylist = [y1]
    while True:
        y1 += distance
        if y1 < y2: 
            ylist.append(y1)
        else:
            break
    
    result = []
    for x3 in xlist:
        for y3 in ylist:
            result.append([x3, y3, 0])
    return result



@acad.decorator_command
def ll_hatch_circle():
    r = acad.GetDouble(5, "请输入圆的半径:")
    distance = acad.GetDouble(10, "请输入圆边与圆边的间隔:")
    distance = 2*r + distance
    objid = acad.EntSel([[0, "REGION"]])
    po1, po2 = acad.GetCorner2()
    meshptlist = __make_rectangle_mesh(po1, po2, distance)
    with acad.transaction() as trans:
        region = acad.TransObjectForWrite(objid)
        brep = acad.Brep(region)
        for po1 in meshptlist:
            pt1 = acad.Vec3Add(po1, [ r,  0])
            pt2 = acad.Vec3Add(po1, [ 0,  r])
            pt3 = acad.Vec3Add(po1, [-r,  0])
            pt4 = acad.Vec3Add(po1, [ 0, -r])
            flag = __ptlist_in_region([po1, pt1, pt2, pt3, pt4], brep)
            if flag: acad.AddCircle(po1, r)







# opts = new PromptSelectionOptions();
# opts.AllowSubSelections = true;
# AllowSubSelections allows to select sub entities (solid or meshes edges, attribute references, ...) accordingly with the LEGACYCTRLPICK and SUBOBJSELECTMODE sysvar settings.
# With default settings (LEGACYCTRLPICK = 2 and SUBOBJSELECTMODE = 0), AllowSubSelections allows sub entities selection using Ctrl+click.
# 你好
# AllowSubSelections 允许通过 LEGACYCTRLPICK 和 SUBOBJSELECTMODE sysvar 设置，相应地选择子实体（实体或网格、边缘、属性引用等）。
# 在默认设置（LEGACYCTRLPICK = 2 和 SUBOBJSELECTMODE = 0）下，允许子实体选择时，可以通过 Ctrl+click 选择子实体。