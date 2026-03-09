from Autodesk.AutoCAD.DatabaseServices import SubentityId, SubentityType, FullSubentityPath, MeshFaceterData, IdMapping, ObjectIdCollection, SubDMesh, Curve, Extents3d, Polyline, Polyline3d, Line, Circle, Poly3dType, DBText, Region, BooleanOperationType
from Autodesk.AutoCAD.BoundaryRepresentation import PointContainment, Brep, Face, BrepEntity
import System


import clr

import acad
import academit

import System


def 命令(): 
    academit.添加命令("llsweep-pl-nj-to-pl-sheet-for", llsweep_pl_nj_to_pl_sheet_for)
    academit.添加命令("llsweep-pl-nj-to-pl-zj-rect-for", llsweep_pl_nj_to_pl_zj_rect_for)
    academit.添加命令("llsweep-subtract-ness-cube-for", llsweep_subtract_ness_cube_for)
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
            dr1 = acad.Direct(mid1, mid2)
            [pt2, pt4] = acad.MatrixRotationPointList(mid1, dr1, 90, [pt1, pt3])
            print([pt1, pt2, pt3, pt4, pt1])
            # acad.AddPolyline3d([pt1, pt2, pt3, pt4, pt1])




