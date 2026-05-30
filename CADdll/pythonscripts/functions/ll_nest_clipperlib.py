

import acad
from Autodesk.AutoCAD.BoundaryRepresentation import PointContainment, Brep, Face, BrepEntity
import academit
import acad

from Autodesk.AutoCAD.DatabaseServices import Line, Arc, ObjectId, Spline, Transaction, OpenMode, BlockTable, BlockTableRecord, BlockReference, LayerTableRecord, ObjectIdCollection, TypedValue, DxfCode, DwgVersion
from Autodesk.AutoCAD.DatabaseServices import Entity, DBPoint, Extents3d, Polyline, Polyline3d, Line, Circle, Poly3dType, DBText, MText, Region, DBObjectCollection, Intersect, Group, Curve, BooleanOperationType
from Autodesk.AutoCAD.Geometry import Point2d, Point3d, Point3dCollection, Matrix3d, Vector2d, Vector3d, DoubleCollection

def 命令():  
    pass
    # academit.添加命令("llnest-mksum", llnest_mksum)  
    # academit.添加命令("llnest-mkdirect", llnest_mkdirect)  
    # academit.添加命令("llnest-mkconnect", llnest_mkconnect)  
    # academit.添加命令("llnest-mkconvexhull", llnest_mkconvexhull)  
    # academit.添加命令("llnest-mkminibound", llnest_mkminibound)  
    # academit.添加命令("llnest-region-to-mkpolyline", llnest_region_to_mkpolyline)  
    # academit.添加命令("llnest-region-intersect", llnest_region_intersect)  


def RECopy(objref):
    matrix4x4 = Matrix3d.Identity # 单位矩阵
    copyobjref = objref.GetTransformedCopy(matrix4x4)
    return copyobjref


@acad.decorator_command
def llnest_region_intersect():
    objidlist = acad.SSGetIdList([[0, "REGION"]])  
    if len(objidlist) < 2: return
    
    with acad.transaction() as trans:
        region1 = trans.GetObject(objidlist[0], OpenMode.ForRead)
        region2 = trans.GetObject(objidlist[1], OpenMode.ForRead)
        if region1.Area < region2.Area: 
            regionmin = region1.Clone()
            regionmax = region2.Clone()
        else:
            regionmin = region2.Clone()
            regionmax = region1.Clone()
            
        regionmin.BooleanOperation(BooleanOperationType.BoolIntersect, regionmax) # 
        flag = False
        if regionmin.Area > 0: # 若无交集，areamin会变为零，areamax则没变化
            flag = True
            acad.AddDBObject(regionmin, "排版1", 2)
    return flag


# public enum BooleanOperationType
# {
#   BoolUnite,
#   BoolIntersect,
#   BoolSubtract,
# }






def point_to_pt1(point):
    return [point.X, point.Y, point.Z]

def pointlist_to_ptlist(pointlist):
    ptlist = []
    for point in pointlist: ptlist.append([point.X, point.Y, point.Z])
    return ptlist

@acad.decorator_command
def llnest_region_to_mkpolyline():
    objid = acad.EntSel([[0, "REGION"]])
    with acad.transaction() as trans:
        objref = trans.GetObject(objid, OpenMode.ForRead)
        brep = Brep(objref)
        for face in brep.Faces: 
            for loop in face.Loops:
                buflist = []
                for edge in loop.Edges: 
                    if edge.Curve.IsLineSegment: 
                        pointlist = curve_linesegment_point(edge)
                    else:
                        pointlist = curve_nurb_point(edge, 5)
                    buflist.append(pointlist)

                # 使有序，region里loop里的每段edge起点和终点是无序的
                count = len(buflist)
                pt2 = point_to_pt1(buflist[0][-1])
                for i in range(1, count):
                    pointlist = buflist[i]
                    po1 = point_to_pt1(pointlist[0])
                    if not acad.IsPointSame(pt2, po1): buflist[i] = pointlist[::-1]
                    pt2 = point_to_pt1(buflist[i][-1])

                cuflist = []
                for pointlist in buflist: cuflist = cuflist + pointlist[0:-1]

                ptlist = pointlist_to_ptlist(cuflist)
                if not acad.IsCCW(ptlist): cuflist = cuflist[::-1]

                polyline = Polyline()
                for i, point in enumerate(cuflist):
                    polyline.AddVertexAt(i, Point2d(point.X, point.Y), 0, 0, 0)
                    acad.AddText([point.X, point.Y], str(i), 1)
                polyline.Closed = True
                acad.AddDBObject(polyline, "排版1", 3)

def curve_linesegment_polyline(edge):
    polyline = Polyline()
    point2d1 = Point2d(edge.Curve.NativeCurve.StartPoint.X, edge.Curve.NativeCurve.StartPoint.Y)
    point2d2 = Point2d(edge.Curve.NativeCurve.EndPoint.X, edge.Curve.NativeCurve.EndPoint.Y)
    polyline.AddVertexAt(0, point2d1, 0, 0, 0)
    polyline.AddVertexAt(1, point2d2, 0, 0, 0)
    return polyline

def curve_nurb_polyline(edge, segment=3):
    nc3d = edge.GetCurveAsNurb()
    dbcurve = Curve.CreateFromGeCurve(nc3d)
    startparam = dbcurve.StartParam
    endparam = dbcurve.EndParam
    distance = dbcurve.GetDistanceAtParameter(endparam)
    # disparam = endparam-startparam
    dparam = segment / distance * (endparam-startparam)
    count = int(distance / segment)
    paramlist = [startparam]
    sumparam = startparam
    for i in range(count):
        sumparam += dparam
        if sumparam >= endparam: break
        paramlist.append(sumparam)
    paramlist.append(endparam)
    pointlist = [dbcurve.GetPointAtParameter(param) for param in paramlist]
    polyline = Polyline()
    for i, point in enumerate(pointlist):
        polyline.AddVertexAt(i, Point2d(point.X, point.Y), 0, 0, 0)
    return polyline

def curve_linesegment_point(edge):
    point1 = Point3d(edge.Curve.NativeCurve.StartPoint.X, edge.Curve.NativeCurve.StartPoint.Y, edge.Curve.NativeCurve.StartPoint.Z)
    point2 = Point3d(edge.Curve.NativeCurve.EndPoint.X, edge.Curve.NativeCurve.EndPoint.Y, edge.Curve.NativeCurve.EndPoint.Z)
    return [point1, point2]

def curve_nurb_point(edge, segment=2):
    nc3d = edge.GetCurveAsNurb()
    dbcurve = Curve.CreateFromGeCurve(nc3d)
    startparam = dbcurve.StartParam
    endparam = dbcurve.EndParam
    distance = dbcurve.GetDistanceAtParameter(endparam)
    disparam = endparam-startparam
    dparam = segment / distance * disparam
    count = int(distance / segment)
    if count < 8: 
        count = 8
        dparam = disparam / 8
    pointlist = [dbcurve.StartPoint]
    sumparam = startparam
    for i in range(count):
        sumparam += dparam
        if sumparam >= endparam: break    
        point = dbcurve.GetPointAtParameter(sumparam)
        pointlist.append(point)
    pointlist.append(dbcurve.EndPoint)
    return pointlist

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
def llnest_mksum():
    objidlist = acad.SSGetIdList()  
    if objidlist == None: return
    ptlist0 = acad.GetMKPolyLinePointList(objidlist[0])
    ptlist1 = acad.GetMKPolyLinePointList(objidlist[1])
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
        for ptlist3, ptlist4 in buflist:
            if ptlist3 != []: acad.AddMKPolyLine(ptlist3, "排版1", 1)
            if ptlist4 != []: acad.AddMKPolyLine(ptlist4, "排版1", 2)

@acad.decorator_command
def llnest_mkdirect():
    objidlist = acad.SSGetIdList()  
    if objidlist == None: return
    ptlist0 = acad.GetMKPolyLinePointList(objidlist[0])
    ptlist1 = acad.GetMKPolyLinePointList(objidlist[1])
    x1, y1, z1 = ptlist1[0]
    drm = [-x1, -y1]
    ptlist1 = [acad.Vec3Add(po1, drm) for po1 in ptlist1]
    buflist = []
    for pt1 in ptlist0:
        ptlist3 = []
        for x1, y1, z1 in ptlist1:
            pt3 = acad.Vec3Add(pt1, [-x1, -y1])
            ptlist3.append(pt3)
        buflist.append([ptlist3, ptlist4])
    with acad.transaction() as trans:
        for ptlist3, ptlist4 in buflist:
            if ptlist3 != []: acad.AddMKPolyLine(ptlist3, "排版1", 1)



@acad.decorator_command
def llnest_mkconnect():
    objidlist = acad.SSGetIdList()  
    if objidlist == None: return
    ptlist0 = acad.GetMKPolyLinePointList(objidlist[0])
    ptlist1 = acad.GetMKPolyLinePointList(objidlist[1])
    x1, y1, z1 = ptlist1[0]
    drm = [-x1, -y1]
    ptlist1 = [acad.Vec3Add(po1, drm) for po1 in ptlist1]
    ptlist1 = ptlist1[1:]
    buflist = []
    for x1, y1, z1 in ptlist1:
        ptlist3 = []
        ptlist4 = []
        for pt1 in ptlist0:
            pt4 = acad.Vec3Add(pt1, [-x1, -y1])
            ptlist4.append(pt4)
        buflist.append(ptlist4)
    with acad.transaction() as trans:
        for i, ptlist3 in enumerate(buflist):
            if ptlist3 != []: acad.AddMKPolyLine(ptlist3, "排版1", i+1)


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




def build_nest_polyline(objid):
    with acad.transaction() as trans:
        objref = trans.GetObject(objid, OpenMode.ForRead)
        match str(objref): 
            case "Autodesk.AutoCAD.DatabaseServices.Polyline": pass
            case "Autodesk.AutoCAD.DatabaseServices.BlockReference": pass
            case "Autodesk.AutoCAD.DatabaseServices.Line": pass


def build_nest_region(objidlist):
    with acad.transaction() as trans:
        objreflist = []
        for objid in objidlist:
            objref = trans.GetObject(objid, OpenMode.ForRead)
            if objref.Layer != "0": continue
            matrix4x4 = Matrix3d.Identity # 单位矩阵
            objrefcopy = objref.GetTransformedCopy(matrix4x4)
            objreflist.append(objrefcopy)
        
    collect = DBObjectCollection()
    for objref in objreflist:
        collect.Add(objref)

    regions = Region.CreateFromCurves(collect)
    buflist = []
    for region in regions:
        area = region.Area
        buflist.append([area, region])
    buflist.sort(key = lambda item: item[0], reverse=True) # 排序规则，reverse = True 降序， reverse = False 升序（默认）

    result = [region for area, region in buflist]
    return result






# 以下是一个射线法的Python实现：
def is_point_in_polygon(ptlist, mkptlist):
    if acad.IsPointSame(mkptlist[0], mkptlist[-1]): 
        lwptlist = mkptlist
    else:
        lwptlist = mkptlist[:]+mkptlist[0:1]

    flag = True
    for pt1 in ptlist:
        inside = is_pt1_in_lwptlist(pt1, lwptlist)
        if inside != True:
            flag = False
            break
    return flag

def is_pt1_in_lwptlist(pt1, lwptlist):
    x, y = pt1[0:2]
    inside = False
    p1x, p1y = lwptlist[0][0:2]
    for pt2 in lwptlist:
        p2x, p2y = pt2[0:2]
        if y > min(p1y, p2y):
            if y <= max(p1y, p2y):
                if x <= max(p1x, p2x):
                    if abs(p1y-p2y) > 0.00001: # p1y != p2y
                        xinters = (y - p1y) * (p2x - p1x) / (p2y - p1y) + p1x
                    if abs(p1x-p2x) < 0.00001 or x <= xinters: # p1y == p2y
                        inside = not inside
        p1x, p1y = p2x, p2y
    return inside

def is_pt1_in_region(pt1, region):
    inside = False
    brep = Brep(region)
    result = brep.GetPointContainment(acad.ToPoint3d(pt1), acad.PointContainment.Outside) # out ref 被自动处理成结果返回
    if result[1] == acad.PointContainment.OnBoundary: # (None, <PointContainment.Outside: 1>)
        inside = True
    return inside

def is_point_in_region(ptlist, region): 
    flag = True
    for pt1 in ptlist:
            inside = is_pt1_in_region(pt1, region)
            if inside != True:
                flag = False
                break
    return flag








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