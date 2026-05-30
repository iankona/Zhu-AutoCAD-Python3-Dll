

import acad
from Autodesk.AutoCAD.BoundaryRepresentation import PointContainment, Brep, Face, BrepEntity
import academit
import acad
import System




from Autodesk.AutoCAD.DatabaseServices import Line, Arc, ObjectId, Spline, Transaction, OpenMode, BlockTable, BlockTableRecord, BlockReference, LayerTableRecord, ObjectIdCollection, TypedValue, DxfCode, DwgVersion
from Autodesk.AutoCAD.DatabaseServices import Entity, DBPoint, Extents3d, Polyline, Polyline3d, Line, Circle, Poly3dType, DBText, MText, Region, DBObjectCollection, Intersect, Group, Curve, BooleanOperationType
from Autodesk.AutoCAD.Geometry import Point2d, Point3d, Point3dCollection, Matrix3d, Vector2d, Vector3d, DoubleCollection

def 命令():  
    academit.添加命令("llnest", llnest)  
    academit.添加命令("llnest-nfp", llnest_nfp)  
    academit.添加命令("llnest-mksum", llnest_mksum)  
    academit.添加命令("llnest-guijixian", llnest_guijixian)  
    academit.添加命令("llnest-mksum-and-guijixian", llnest_mksum_and_guijixian)
    academit.添加命令("llnest-mkconvexhull", llnest_mkconvexhull)  
    academit.添加命令("llnest-mkminibound", llnest_mkminibound)  



def acad_add_point_text(ptlist):
    with acad.transaction() as trans:
        for i, pt1 in enumerate(ptlist):
            acad.AddText(pt1, str(i), 30)

def acad_pointlist_move_zero(ptlist):
    x0, y0, z0 = ptlist[0]
    dr0 = [-x0, -y0, -z0]
    result = []
    for pt1 in ptlist:
        pt1 = acad.Vec3Add(pt1, dr0)
        result.append(pt1)
    return result

            
def acad_pointlist_add_direct(ptlist, dr1):
    result = []
    for pt1 in ptlist:
        pt1 = acad.Vec3Add(pt1, dr1)
        result.append(pt1)
    return result


@acad.decorator_command
def llnest_mksum():
    objidlist = acad.SSGetIdList()  
    if len(objidlist) < 2: return
    po1 = acad.GetPoint()
    if po1 == None: return 
    ptlist0 = acad.GetMKPolyLinePointList(objidlist[0])
    ptlist1 = acad.GetMKPolyLinePointList(objidlist[1])
    acad_add_point_text(ptlist0)
    acad_add_point_text(ptlist1)
    ptlist0 = acad_pointlist_move_zero(ptlist0)
    ptlist1 = acad_pointlist_move_zero(ptlist1)

    buflist = []
    for pt1 in ptlist0:
        ptlist3 = []
        ptlist4 = []
        for [x1, y1, z1] in ptlist1:
            pt3 = acad.Vec3Add(pt1, [ x1,  y1])
            pt4 = acad.Vec3Add(pt1, [-x1, -y1])
            ptlist3.append(pt3)
            ptlist4.append(pt4)
        buflist.append([ptlist3, ptlist4])
    with acad.transaction() as trans:
        ptlist0 = acad_pointlist_add_direct(ptlist0, po1)
        # ptlist1 = acad_pointlist_add_direct(ptlist1, po1)
        acad.AddMKPolyLine(ptlist0, "排版1", 0)
        for ptlist3, ptlist4 in buflist:
            ptlist3 = acad_pointlist_add_direct(ptlist3, po1)
            ptlist4 = acad_pointlist_add_direct(ptlist4, po1)
            acad_add_point_text(ptlist3)
            acad_add_point_text(ptlist4)
            if ptlist3 != []: acad.AddMKPolyLine(ptlist3, "排版1", 1)
            if ptlist4 != []: acad.AddMKPolyLine(ptlist4, "排版1", 2)

# @acad.decorator_command
# def llnest_mkdirect():
#     objidlist = acad.SSGetIdList()  
#     if objidlist == None: return
#     ptlist0 = acad.GetMKPolyLinePointList(objidlist[0])
#     ptlist1 = acad.GetMKPolyLinePointList(objidlist[1])
#     x1, y1, z1 = ptlist1[0]
#     drm = [-x1, -y1]
#     ptlist1 = [acad.Vec3Add(po1, drm) for po1 in ptlist1]
#     buflist = []
#     for pt1 in ptlist0:
#         ptlist3 = []
#         for x1, y1, z1 in ptlist1:
#             pt3 = acad.Vec3Add(pt1, [-x1, -y1])
#             ptlist3.append(pt3)
#         buflist.append([ptlist3, ptlist4])
#     with acad.transaction() as trans:
#         for ptlist3, ptlist4 in buflist:
#             if ptlist3 != []: acad.AddMKPolyLine(ptlist3, "排版1", 1)



@acad.decorator_command
def llnest_guijixian():
    objidlist = acad.SSGetIdList()  
    if len(objidlist) < 2: return
    po1 = acad.GetPoint()
    if po1 == None: return 
    ptlist0 = acad.GetMKPolyLinePointList(objidlist[0])
    ptlist1 = acad.GetMKPolyLinePointList(objidlist[1])
    acad_add_point_text(ptlist0)
    acad_add_point_text(ptlist1)
    ptlist0 = acad_pointlist_move_zero(ptlist0)
    ptlist1 = acad_pointlist_move_zero(ptlist1)

    buflist = []

    for [x1, y1, z1] in ptlist1:
        ptlist3 = []
        ptlist4 = []
        for pt1 in ptlist0:
            pt3 = acad.Vec3Add(pt1, [ x1,  y1])
            pt4 = acad.Vec3Add(pt1, [-x1, -y1])
            ptlist3.append(pt3)
            ptlist4.append(pt4)
        buflist.append([ptlist3, ptlist4])

    with acad.transaction() as trans:
        ptlist0 = acad_pointlist_add_direct(ptlist0, po1)
        # ptlist1 = acad_pointlist_add_direct(ptlist1, po1)
        acad.AddMKPolyLine(ptlist0, "排版1", 0)
        for ptlist3, ptlist4 in buflist:
            ptlist3 = acad_pointlist_add_direct(ptlist3, po1)
            ptlist4 = acad_pointlist_add_direct(ptlist4, po1)
            acad_add_point_text(ptlist3)
            acad_add_point_text(ptlist4)
            if ptlist3 != []: acad.AddMKPolyLine(ptlist3, "排版1", 1)
            if ptlist4 != []: acad.AddMKPolyLine(ptlist4, "排版1", 2)



@acad.decorator_command
def llnest_mksum_and_guijixian():
    objidlist = acad.SSGetIdList()  
    if len(objidlist) < 2: return
    po1 = acad.GetPoint()
    if po1 == None: return 
    ptlist0 = acad.GetMKPolyLinePointList(objidlist[0])
    ptlist1 = acad.GetMKPolyLinePointList(objidlist[1])
    # acad_add_point_text(ptlist0)
    # acad_add_point_text(ptlist1)
    ptlist0 = acad_pointlist_move_zero(ptlist0)
    ptlist1 = acad_pointlist_move_zero(ptlist1)

    buflist = []

    # mksum
    for pt1 in ptlist0:
        ptlist3 = []
        ptlist4 = []
        for [x1, y1, z1] in ptlist1:
            pt3 = acad.Vec3Add(pt1, [ x1,  y1])
            pt4 = acad.Vec3Add(pt1, [-x1, -y1])
            ptlist3.append(pt3)
            ptlist4.append(pt4)
        buflist.append([ptlist3, ptlist4])
    
    #轨迹线
    for [x1, y1, z1] in ptlist1:
        ptlist3 = []
        ptlist4 = []
        for pt1 in ptlist0:
            pt3 = acad.Vec3Add(pt1, [ x1,  y1])
            pt4 = acad.Vec3Add(pt1, [-x1, -y1])
            ptlist3.append(pt3)
            ptlist4.append(pt4)
        buflist.append([ptlist3, ptlist4])

    with acad.transaction() as trans:
        for ptlist3, ptlist4 in buflist:
            ptlist3 = acad_pointlist_add_direct(ptlist3, po1)
            ptlist4 = acad_pointlist_add_direct(ptlist4, po1)
            # acad_add_point_text(ptlist3)
            # acad_add_point_text(ptlist4)
            # if ptlist3 != []: acad.AddMKPolyLine(ptlist3, "排版1", 1)
            if ptlist4 != []: acad.AddMKPolyLine(ptlist4, "排版1", 2)








def acad_break_find_objreflist_acadpointlist(break_objreflist):
    buflist = []
    for i, objref1 in enumerate(break_objreflist):
        acad_point_list = []
        for j, objref2 in enumerate(break_objreflist):
            if i == j: continue
            # print([str(objref1.ObjectId),str(objref2.ObjectId)]) # ['(0)', '(0)'] # 没有trans到cad里是没有objid的
            # if str(objref1.ObjectId) == str(objref2.ObjectId): continue
            collect = acad.Point3dCollection() 
            objref1.IntersectWith(objref2, acad.Intersect.OnBothOperands, collect, System.IntPtr.Zero, System.IntPtr.Zero)
            for point in collect: 
                acad_point_list.append(point)
        buflist.append([objref1, acad_point_list])
    return buflist

def acad_break_objref_with_acadpointlist(objref, acadpointlist):
    pt1, pt2 = [objref.StartPoint.X, objref.StartPoint.Y, objref.StartPoint.Z], [objref.EndPoint.X, objref.EndPoint.Y, objref.EndPoint.Z]
    para_list = []
    for point in acadpointlist:
        po1 = [point.X, point.Y, point.Z]
        if acad.IsPointSame(pt1, po1): continue
        if acad.IsPointSame(pt2, po1): continue
        para = objref.GetParameterAtPoint(point)
        para_list.append(para)
    if para_list == []: return []
    para_list.sort() # 默认从小到大
    collect = acad.DoubleCollection()
    for para in para_list: collect.Add(para)
    result = objref.GetSplitCurves(collect)
    buflist = []
    for subobjref in result:
        buflist.append(subobjref)
    return buflist


def acad_explode(objreflist):
    hasexplode = False
    redsultref = []
    for objref in objreflist:
        match str(objref): 
            case "Autodesk.AutoCAD.DatabaseServices.Polyline": 
                result = acad.DBObjectCollection()
                objref.Explode(result)
                for subobjref in result: redsultref.append(subobjref)
                hasexplode = True
            case "Autodesk.AutoCAD.DatabaseServices.BlockReference": 
                result = acad.DBObjectCollection()
                objref.Explode(result)
                for subobjref in result: redsultref.append(subobjref)
                hasexplode = True
            case _: 
                redsultref.append(objref)
    return hasexplode, redsultref



def acad_union_region(regionlist):
    regionlist.sort(key = lambda item: item[0], reverse=True) # 排序规则，reverse = True 降序， reverse = False 升序（默认）
    reginmax = regionlist[0][-1]
    for eare, region in regionlist[1:]:
        reginmax.BooleanOperation(BooleanOperationType.BoolUnite, region)
    return reginmax


# regionmin.BooleanOperation(BooleanOperationType.BoolIntersect, regionmax) # 
# flag = False
# if regionmin.Area > 0: # 若无交集，areamin会变为零，areamax则没变化
#     flag = True
#     acad.AddDBObject(regionmin, "排版1", 2)
# return flag


# public enum BooleanOperationType
# {
#   BoolUnite,
#   BoolIntersect,
#   BoolSubtract,
# }





def acad_pointlist_to_edgelist(ptlist):
    edgelist = []
    count = len(ptlist)
    for i in range(count):
        j = (i+1) % count
        pt1 = ptlist[i]
        pt2 = ptlist[j]
        edgelist.append([pt1, pt2])
    return edgelist
    

def acad_ptlist2_to_regionmax(ptlista, ptlistb):
    objreflist = []
    for pa1, pb1 in zip(ptlista, ptlistb):
        line = acad.DBObjectLine(pa1, pb1)
        objreflist.append(line)
    objreflist.append(acad.DBObjectMKPolyLine(ptlista))
    objreflist.append(acad.DBObjectMKPolyLine(ptlistb))

    # break all
    buflist = acad_break_find_objreflist_acadpointlist(objreflist) 
    cuflist = []
    for objref, acad_point_list in buflist:
        result = acad_break_objref_with_acadpointlist(objref, acad_point_list)
        if result == []:
            cuflist.append(objref)
        else:
            for subobjref in result: 
                cuflist.append(subobjref)
    
    while True:
        hasexplode, cuflist = acad_explode(cuflist)
        if hasexplode == False: break

    collect = acad.DBObjectCollection()
    for objref in cuflist: collect.Add(objref)

    regions = Region.CreateFromCurves(collect)
    auflist = []
    for region in regions:
        area = region.Area
        auflist.append([area, region])
    auflist.sort(key = lambda item: item[0], reverse=True) # 排序规则，reverse = True 降序， reverse = False 升序（默认）
    return auflist[0]

def nest_mincowsky_sum(ptlist0, ptlist1, pb1):
    edgelist = acad_pointlist_to_edgelist(ptlist0)


    # mksum
    ptlist1 = [ [x1, y1, 0] for [x1, y1, z1] in ptlist1]

    buflist = []
    for po1, po2 in edgelist:
        po1 = acad.Vec3Add(po1, pb1)
        po2 = acad.Vec3Add(po2, pb1)
        ptlista = acad_pointlist_add_direct(ptlist1, po1)
        ptlistb = acad_pointlist_add_direct(ptlist1, po2)
        buflist.append([ptlista, ptlistb])

    regionlist = []
    for ptlista, ptlistb in buflist:
        area, region = acad_ptlist2_to_regionmax(ptlista, ptlistb)
        regionlist.append([area, region])

    return regionlist


def nest_mincowsky_diff(ptlist0, ptlist1, pb1=None):
    if pb1 != None:
        baseptlist = acad_pointlist_add_direct(ptlist0, pb1)
    else:
        baseptlist = ptlist0
    objref = acad.DBObjectMKPolyLine(baseptlist)
    collect = acad.DBObjectCollection()
    collect.Add(objref)
    regions = Region.CreateFromCurves(collect)
    baseregion = regions[0]

    edgelist = acad_pointlist_to_edgelist(ptlist0)
    # mkdiff
    ptlist1 = [ [-x1, -y1, 0] for [x1, y1, z1] in ptlist1]

    buflist = []
    for po1, po2 in edgelist:
        if pb1 != None:
            po1 = acad.Vec3Add(po1, pb1)
            po2 = acad.Vec3Add(po2, pb1)
        ptlista = acad_pointlist_add_direct(ptlist1, po1)
        ptlistb = acad_pointlist_add_direct(ptlist1, po2)
        buflist.append([ptlista, ptlistb])

    regionlist = []
    for ptlista, ptlistb in buflist:
        area, region = acad_ptlist2_to_regionmax(ptlista, ptlistb)
        regionlist.append([area, region])

    return baseregion, regionlist





def nest_nfp_region_to_sidelist(regionnfp, baseregion):
    brep = Brep(regionnfp)
    looplist = []
    for face in brep.Faces:
        for loop in face.Loops:
            edgelist = []
            for edge in loop.Edges:
                point1 = edge.Vertex1.Point
                point2 = edge.Vertex2.Point
                pt1, pt2 = [point1.X, point1.Y, point1.Z], [point2.X, point2.Y, point2.Z]
                edgelist.append([pt1, pt2])
            looplist.append(edgelist)

    collect = acad.DBObjectCollection()
    for loop in looplist:
        for pt1, pt2 in loop:
            objref = acad.DBObjectLine(pt1, pt2)
            collect.Add(objref)

    regions = Region.CreateFromCurves(collect)

    auflist = []
    for region in regions:
        area = region.Area
        auflist.append([area, region])
    auflist.sort(key = lambda item: item[0], reverse=True) # 排序规则，reverse = True 降序， reverse = False 升序（默认）

    nfp = auflist[0][-1]
    outside = [nfp]
    inside = []
    for area, region in auflist[1:]:
        regionmin = region.Clone()
        regionmax = baseregion.Clone()
        regionmin.BooleanOperation(BooleanOperationType.BoolIntersect, regionmax)
        if regionmin.Area <= 0: outside.append(region)
        if regionmin.Area >= region.Area: inside.append(region)
    return outside, inside








@acad.decorator_command
def llnest_nfp():
    objidlist = acad.SSGetIdList()  
    if len(objidlist) < 2: return
    pb1 = acad.GetPoint()
    if pb1 == None: return 
    ptlist0 = acad.GetMKPolyLinePointList(objidlist[0])
    ptlist1 = acad.GetMKPolyLinePointList(objidlist[1])
    ptlist0 = acad_pointlist_move_zero(ptlist0)
    ptlist1 = acad_pointlist_move_zero(ptlist1)
    baseregion, regionlist = nest_mincowsky_diff(ptlist0, ptlist1, pb1)
    regionmax = acad_union_region(regionlist)
    outsidelist, insidelist = nest_nfp_region_to_sidelist(regionmax, baseregion)
    
    # with acad.transaction() as trans:
    #     for area, region in regionlist:
    #         acad.AddDBObject(region, "排版1", 1)

    # with acad.transaction() as trans:
    #     acad.AddDBObject(regionmax, "排版1", 1)

    with acad.transaction() as trans:
        acad.AddDBObject(baseregion, "排版1", 0)
        for region in outsidelist:
            acad.AddDBObject(region, "排版1", 1)
        for region in insidelist:
            acad.AddDBObject(region, "排版1", 2)








def acad_idlist_to_reflist(objidlist):
    with acad.transaction() as trans:
        objreflist = []
        for objid in objidlist:
            objref = trans.GetObject(objid, OpenMode.ForRead)
            objreflist.append(objref)
    return objreflist

def acad_copy_reflist(objreflist):
    objrefcopylist = []
    for objref in objreflist:
        matrix4x4 = Matrix3d.Identity # 单位矩阵
        objrefcopy = objref.GetTransformedCopy(matrix4x4)
        objrefcopylist.append(objrefcopy)
    return objrefcopylist

def acad_add_blockid(objreflist, basepoint):
    with acad.transaction() as trans:
        pt1, pt2 = acad.GetRefListBoundXY0(objreflist)
        blockobjid = acad.DBObjectBlockFromRefList(objreflist, basepoint)
        # pt1 = acad.Vec3Add(pt1, [1500,0,0])
        # objref = acad.DBObjectBlockRef(blockobjid, pt1)
        # acad.DBObjectMove(objref, [0,0,0], [1500,0,0])
        # acad.AddDBObject(objref)
    return blockobjid





def acad_add_regionlist(regionlist):
    with acad.transaction() as trans:
        objidlist = []
        for region in regionlist:
            acad.DBObjectMove(region, [0,0,0], [3000,0,0])
            objid = acad.AddDBObject(region)
            objidlist.append(region.ObjectId)
        acad.AddGroup(objidlist)


def acad_find_region_and_hold(objidlist):
    objreflist = acad_idlist_to_reflist(objidlist)
    objrefcopylist = acad_copy_reflist(objreflist)
    while True:
        hasexplode, objrefcopylist = acad_explode(objrefcopylist)
        if hasexplode == False: break
    objrefcopylist1 = objrefcopylist
    objrefcopylist2 = acad_copy_reflist(objrefcopylist)

    blockobjreflist = []
    for objref in objrefcopylist1:
        if "Dimension" in str(objref): continue
        if objref.Layer == "0": blockobjreflist.append(objref)
        if objref.Layer == "打标1": blockobjreflist.append(objref)
    
    regionobjreflist = []
    for objref in objrefcopylist2:
        if "Dimension" in str(objref): continue
        if objref.Layer == "0": regionobjreflist.append(objref)

    collect = DBObjectCollection()
    for objref in regionobjreflist: collect.Add(objref)

    regions = Region.CreateFromCurves(collect)
    buflist = []
    for region in regions:
        area = region.Area
        buflist.append([area, region])
    buflist.sort(key = lambda item: item[0], reverse=True) # 排序规则，reverse = True 降序， reverse = False 升序（默认）
    regionlist = [region for area, region in buflist]
    return blockobjreflist, regionlist


def acad_find_block_count(blockobjreflist):
    count = 1
    for objref in blockobjreflist:
        if str(objref) == "Autodesk.AutoCAD.DatabaseServices.DBText":
            string = objref.TextString
            if "件" in string: 
                count = int(string[0:-1])
                break
            # if "x" in string or "X" in string: # Error 3000x1530
            #     count = int(string[1:])
            #     break
    return count



class AcadPart:
    def __init__(self):
        pass

    def process(self, objidlist=[]):
        self.objidlist = objidlist
        objreflist, regionlist = acad_find_region_and_hold(objidlist)
        self.region = regionlist[0] 
        self.holdlist = regionlist[1:]
        self.basepoint = nest_get_region_pb0(self.region)
        self.blockid = acad_add_blockid(objreflist, self.basepoint)
        self.blockcount = acad_find_block_count(objreflist)
        # self.targetposition = [0,0,0]
        # self.targetrotation = 0
        # acad_add_regionlist(regionlist)
        return self

    def copy(self):
        part = AcadPart()
        part.objidlist = self.objidlist
        part.region = acad.DBObjectCopy(self.region)
        part.holdlist = [acad.DBObjectCopy(region) for region in self.holdlist]
        part.basepoint = self.basepoint
        part.blockid = self.blockid
        part.blockcount = self.blockcount
        # part.targetposition = [0,0,0]
        # part.targetrotation = 0
        return part







def acadpartlist_to_nestpartlist(acadpartlist):
    nestpartlist = []
    for part in acadpartlist:
        for i in range(part.blockcount): # 已经默认为1
            nestpartlist.append(part.copy())
    return nestpartlist



def nest_calc_nfp(ptlist0, ptlist1): # ptlist0 固定
    # ptlist0 = acad_pointlist_move_zero(ptlist0) 
    ptlist1 = acad_pointlist_move_zero(ptlist1)
    baseregion, regionlist = nest_mincowsky_diff(ptlist0, ptlist1)
    regionmax = acad_union_region(regionlist)
    outsidelist, insidelist = nest_nfp_region_to_sidelist(regionmax, baseregion)

    # with acad.transaction() as trans:
    #     acad.AddDBObject(baseregion, "排版1", 0)
    #     for region in outsidelist:
    #         acad.AddDBObject(region, "排版1", 1)
    #     for region in insidelist:
    #         acad.AddDBObject(region, "排版1", 2)

    return outsidelist, insidelist


def nest_region_to_mkptlist(region):
    brep = Brep(region)
    edgelist = []
    for edge in brep.Edges:
        point1 = edge.Vertex1.Point
        point2 = edge.Vertex2.Point
        pt1 = [point1.X, point1.Y, point1.Z]
        pt2 = [point2.X, point2.Y, point2.Z]
        edgelist.append([pt1, pt2])

    edgelist = loopsort(edgelist)
    ptlist = []
    for pt1, pt2 in edgelist: ptlist.append(pt1)
    return ptlist
    
def nest_get_region_pb0(region):
    brep = Brep(region)
    for edge in brep.Edges:
        point1 = edge.Vertex1.Point
        pt1 = [point1.X, point1.Y, point1.Z]
        return pt1




def nest_region_get_espoint(region):
    brep = Brep(region)
    ptlist = []
    for face in brep.Faces:
        for loop in face.Loops: 
            for edge in loop.Edges: # autocad loop 允许无序 ... ...
                point1 = edge.Vertex1.Point
                point2 = edge.Vertex2.Point
                pt1 = [point1.X, point1.Y, point1.Z]
                pt2 = [point2.X, point2.Y, point2.Z]
                ptlist.append(pt1)
                ptlist.append(pt2)

    x0, y0, z0 = ptlist[0]
    ymin = y0
    for x1, y1, z1 in ptlist[1:]:
        if y1 < ymin: ymin = y1

    ptminlist = []
    for x1, y1, z1 in ptlist:
        if abs(y1-ymin) < 0.000015: ptminlist.append([x1, y1, z1])

    x0, y0, z0 = ptminlist[0]
    xmax = x0
    for x1, y1, z1 in ptminlist[1:]:
        if x1 > xmax: xmax = x1

    ptes = None
    for x1, y1, z1 in ptminlist:
        if abs(x1-xmax) < 0.000015: ptes = [x1, y1, z1]
    return ptes


class Nest:
    def __init__(self):
        pass

    def setacadpart(self, acadpartlist=[], lwcontainlist=[]):
        self.acadpartlist = acadpartlist
        self.nestpartlist = acadpartlist_to_nestpartlist(acadpartlist)
        self.lwcontainlist = lwcontainlist
        self.layoutpartlist = []
        return self

    def layoutpart(self):
        pass
        # ptlist0 = self.lwcontainlist[0]
        # region = self.nestpartlist[0].region
        # ptlist1 = nest_region_to_mkptlist(region)
        # outsidelist, insidelist = nest_calc_nfp(ptlist0, ptlist1)



        for npart in self.nestpartlist:
            ptlist0 = self.lwcontainlist[0]
            ptlist1 = nest_region_to_mkptlist(npart.region)
            a, insidelist = nest_calc_nfp(ptlist0, ptlist1)
            pb1 = nest_get_region_pb0(npart.region)
            buflist = []
            for lpart in self.layoutpartlist:
                ptlist0 = nest_region_to_mkptlist(lpart.region) 
                outsidelist, a = nest_calc_nfp(ptlist0, ptlist1)
                for region in outsidelist: buflist.append(region)

            # print(insidelist, buflist)
            region0 = insidelist[0]
            # with acad.transaction() as trans:
            #     acad.AddDBObject(region0.Clone())

            for region1 in buflist:
                # print(region0.Area, region1.Area)
                region0.BooleanOperation(BooleanOperationType.BoolSubtract, region1)
                # with acad.transaction() as trans:
                #     acad.AddDBObject(region1.Clone())
                # print(region0.Area, region1.Area)

            
            with acad.transaction() as trans:
                pb2 = nest_region_get_espoint(region0)
                # acad.AddDBObject(npart.region.Clone(),color_index=1)
                acad.DBObjectMove(npart.region, pb1, pb2)
                # acad.AddDBObject(region0)
                # acad.AddDBObject(npart.region.Clone(),color_index=2)
                objref = acad.DBObjectBlockRef(npart.blockid, pb2)
                acad.AddDBObject(objref)

            self.layoutpartlist.append(npart)



@acad.decorator_command
def llnest():
    objidlist = acad.SSGetIdList()  
    objid0 = acad.EntSel([[0, "LWPOLYLINE"]])
    if len(objidlist) < 2: return
    pb1 = acad.GetPoint()
    if pb1 == None: return 
    with acad.transaction() as trans:
        lwptlist = acad.TransLWPolyLinePointList(objid0)
        result = acad.TransAutoFindRegionRectList(objidlist)
        partlist = []
        for pt1, pt2 in result: 
            objidlist = acad.GetSelectCornerCrossIdList(pt1, pt2)
            part = AcadPart().process(objidlist)
            partlist.append(part)

    worker = Nest().setacadpart(partlist, [lwptlist])
    worker.layoutpart()

    











def convex_hull_n3(ptlist):
    count = len(ptlist)
    indexlist = []
    result = []
    for i in range(0, count-1):
        for k in range(i, count):
            pt1 = ptlist[i]
            pt2 = ptlist[k]
            flaglist = []
            for n in range(count):
                if n == i: continue
                if n == k: continue
                pt3 = ptlist[n]
                perflag = acad.GetPerflagXY(pt1, pt2, pt3)
                flaglist.append(perflag)
            if 0 in flaglist: continue
            flagset = set(flaglist)
            if len(flagset) == 1:
                result.append([pt1, pt2])
                indexlist.append(i)
                indexlist.append(k)
    return result

def loopsort(listpoint2):
    count  = len(listpoint2)
    buflist = listpoint2[0:1]
    reflist = listpoint2[1:]
    indexlist = []
    for i in range(count-1):
        pt1, pt2 = buflist[-1]
        for k, [po1, po2] in enumerate(reflist):
            if k in indexlist: continue
            flag = False
            if acad.IsPointSame(pt2, po1):
                flag = True
                buflist.append([po1, po2])
            if acad.IsPointSame(pt2, po2):
                flag = True
                buflist.append([po2, po1])
            if flag:
                indexlist.append(k)
                break
    return buflist

def calc_bound_from_ptlist(ptlist):
    x_min, y_min, z_min = x_max, y_max, z_max = ptlist[0]
    for x1, y1, z1 in ptlist[1:]:
        if x1 < x_min: x_min = x1
        if x1 > x_max: x_max = x1
        if y1 < y_min: y_min = y1
        if y1 > y_max: y_max = y1
        # if z1 < z_min: z_min = z1
        # if z1 > z_max: z_max = z1
    return [x_min, y_min, 0], [x_max, y_min, 0], [x_max, y_max, 0], [x_min, y_max, 0], x_max-x_min, y_max-y_min

def minibound_from_listpoint2(listpoint2):
    ptlist0 = [pt1 for pt1, pt2 in listpoint2]
    result = []
    for pt1, pt2 in listpoint2:
        pt0 = pt1
        dr1 = [1,0,0]
        dr2 = acad.Direct(pt1, pt2)
        angle = acad.AngleFromDotDr1Dr2(dr1, dr2)
        axis = acad.CrossNormalized(dr1, dr2)
        ptlist1 = acad.MatrixRotationPointList(angle, axis, pt0, ptlist0)
        pt1, pt2, pt3, pt4, length, width = calc_bound_from_ptlist(ptlist1)
        po1, po2, po3, po4 = acad.MatrixRotationPointList(-angle, axis, pt0, [pt1, pt2, pt3, pt4])
        result.append([length*width, po1, po2, po3, po4])
    result.sort(key = lambda item: item[0], reverse=False) # 从小到大
    return result[0][1:]





@acad.decorator_command
def llnest_mkconvexhull():
    objidlist = acad.SSGetIdList()  
    if objidlist == None: return
    ptlist0 = acad.GetMKPolyLinePointList(objidlist[0])
    ptlist1 = acad.GetMKPolyLinePointList(objidlist[1])
    x1, y1, z1 = ptlist1[0]
    drm = [-x1, -y1]
    ptlist1 = [acad.Vec3Add(po1, drm) for po1 in ptlist1]
    buflist = []
    for x1, y1, z1 in ptlist1:
        for pt1 in ptlist0:
            pt4 = acad.Vec3Add(pt1, [-x1, -y1])
            buflist.append(pt4)
    result = convex_hull_n3(buflist)
    with acad.transaction() as trans:
        pline = acad.TransAutoPt1pt2ListToMKPolyLine(result, "排版1", 1)









@acad.decorator_command
def llnest_mkminibound():
    objidlist = acad.SSGetIdList()  
    if objidlist == None: return
    ptlist0 = acad.GetMKPolyLinePointList(objidlist[0])
    ptlist1 = acad.GetMKPolyLinePointList(objidlist[1])
    x1, y1, z1 = ptlist1[0]
    drm = [-x1, -y1]
    ptlist1 = [acad.Vec3Add(po1, drm) for po1 in ptlist1]
    buflist = []
    for x1, y1, z1 in ptlist1:
        for pt1 in ptlist0:
            pt4 = acad.Vec3Add(pt1, [-x1, -y1])
            buflist.append(pt4)
    result = convex_hull_n3(buflist)
    result = loopsort(result)
    with acad.transaction() as trans:
        ptlist0 = [pt1 for pt1, pt2 in result]
        acad.AddMKPolyLine(ptlist0,  "排版1", 1)
        po1, po2, po3, po4 = minibound_from_listpoint2(result)
        acad.AddMKPolyLine([po1, po2, po3, po4],  "排版1", 2)







# 矩形装箱算法
# def pack(rects):
# rects = sorted(rects, key=lambda r: max(r))
# bins = []
# while rects:
# bin = []
# width_left = BOUND
# height = 0
# for rect in rects:
# if rect[0]<= width_left:
# bin.append(rect)
# width_left -= rect[0]
# height = max(height, rect[1])
# for rect in bin:
# rects.remove(rect)
# bins.append((BOUND - width_left, height, bin))
# return bins
# 以上是一个简单的矩形排料算法实现。代码中首先将矩形按照最大边长从大到小排序。然后依次放入宽度为BOUND的“箱子”中。在每个箱子中，从剩余宽度中选取最大高度的矩形，直到无法再放入为止。
# 这段代码的复杂度为O(N^2)，对于大规模问题可能会比较慢。但对于小规模问题，这种简单的贪心算法已经能够得到较好的结果。