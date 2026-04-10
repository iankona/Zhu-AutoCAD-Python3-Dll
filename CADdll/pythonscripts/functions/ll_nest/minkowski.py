

import acad
from . import convexhull



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
    result = convexhull.convex_hull_n3(buflist)
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
    result = convexhull.convex_hull_n3(buflist)
    result = convexhull.loopsort(result)
    with acad.transaction() as trans:
        ptlist0 = [pt1 for pt1, pt2 in result]
        acad.AddMKPolyLine(ptlist0,  "排版1", 1)
        po1, po2, po3, po4 = convexhull.minibound_from_listpoint2(result)
        acad.AddMKPolyLine([po1, po2, po3, po4],  "排版1", 2)
