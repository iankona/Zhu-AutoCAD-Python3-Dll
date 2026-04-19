
import clr

import acad
import academit

import System

import os
import math
import openpyxl as Excel
import subprocess



def 命令():  
    academit.添加命令("llsheet-to-zj", llsheet_to_zj_for)
    academit.添加命令("llsheet-to-nj", llsheet_to_nj_for)
    academit.添加命令("llsheet-to-bc", llsheet_to_bc_for)
    academit.添加命令("llsheet-to-wj", llsheet_to_wj_for)
    academit.添加命令("llsheet-to-n-rect-for", llsheet_to_n_rect_for)
    academit.添加命令("llsheet_from_nj_for", llsheet_from_nj_for)
    academit.添加命令("llsheet_from_wj_for", llsheet_from_wj_for)

@acad.decorator_command
def llsheet_to_zj_for():
    po1, po2 = acad.GetPoint2()
    if po1 == None: return
    dr1 = acad.Direct(po1, po2)
    while True:
        objidlist = acad.SSGetIdList([[0, "LWPOLYLINE"]])
        if objidlist == None: return
        with acad.transaction() as trans:
            objid1, objid2 = objidlist[0:2]
            objid3, ness = acad.TransAutoFindMidLWPolyLine(objid1, objid2)
            acad.TransMove(objid3, po1, po2)


def zj_to_nj(lengthlist, ness):
    ness_half = ness / 2
    count = len(lengthlist)
    match count:
        case 0: return []
        case 1: lengthlist[0] = lengthlist[0] - ness
        case _:
            lengthlist[0] = lengthlist[0] - ness_half
            for i in range(1, count-1): 
                lengthlist[i] = lengthlist[i] - ness
            lengthlist[-1] = lengthlist[-1] - ness_half
    return lengthlist

def zj_to_wj(lengthlist, ness):
    ness_half = ness / 2
    count = len(lengthlist)
    match count:
        case 0: return []
        case 1: lengthlist[0] = lengthlist[0] + ness
        case _:
            lengthlist[0] = lengthlist[0] + ness_half
            for i in range(1, count-1): 
                lengthlist[i] = lengthlist[i] + ness
            lengthlist[-1] = lengthlist[-1] + ness_half
    return lengthlist

def nj_to_wj(lengthlist, ness):
    ness_double = ness * 2
    count = len(lengthlist)
    match count:
        case 0: return []
        case 1: lengthlist[0] = lengthlist[0] + ness
        case _:
            lengthlist[0] = lengthlist[0] + ness
            for i in range(1, count-1): 
                lengthlist[i] = lengthlist[i] + ness_double
            lengthlist[-1] = lengthlist[-1] + ness
    return lengthlist

def wj_to_nj(lengthlist, ness):
    ness_double = ness * 2
    count = len(lengthlist)
    match count:
        case 0: return []
        case 1: lengthlist[0] = lengthlist[0] - ness
        case _:
            lengthlist[0] = lengthlist[0] - ness
            for i in range(1, count-1): 
                lengthlist[i] = lengthlist[i] - ness_double
            lengthlist[-1] = lengthlist[-1] - ness
    return lengthlist


def nj_to_bc(lengthlist, anglelist, depth):
    buflist = lengthlist[:]
    cuflist = [None]
    for i, angle in enumerate(anglelist):
        if angle == None: continue
        angle = angle / 2 # tan = sin / cos
        shaping = depth / math.tan(acad.Angle2Rad(angle))
        buflist[i-1] = buflist[i-1] + shaping 
        buflist[i]   = buflist[i]   + shaping 
        cuflist.append(shaping)
    cuflist.append(None)
    return buflist, cuflist


@acad.decorator_command
def llsheet_to_nj_for():
    po1, po2 = acad.GetPoint2()
    if po1 == None: return
    dr0 = acad.Direct(po1, po2)
    while True:
        objidlist = acad.SSGetIdList([[0, "LWPOLYLINE"]])
        if objidlist == None: return
        with acad.transaction() as trans:
            objid1, objid2 = objidlist[0:2]
            ptlist = acad.TransLWPolyLinePointList(objid1, po1)
            drlist = acad.TransLWPolyLineDirectList(objid1, po1)
            lnlist, ness = acad.TransNRectLengthListFromSheet(objid1, objid2, po1)
            pt1 = acad.Vec3Add(ptlist[0], dr0)
            ptlist = [pt1]
            for length, direct in zip(lnlist, drlist):
                dr1 = acad.Vec3ResetLength(direct, length)
                pt1 = acad.Vec3Add(pt1, dr1)
                ptlist.append(pt1)
            acad.AddLWPolyLine(ptlist)

@acad.decorator_command
def llsheet_to_bc_for():
    depth = acad.GetDouble(1.0, string="请输入刨槽深度:")
    po1, po2 = acad.GetPoint2()
    if po1 == None: return
    dr0 = acad.Direct(po1, po2)
    while True:
        objidlist = acad.SSGetIdList([[0, "LWPOLYLINE"]])
        if objidlist == None: return
        with acad.transaction() as trans:
            objid1, objid2 = objidlist[0:2]
            ptlist = acad.TransLWPolyLinePointList(objid1, po1)
            drlist = acad.TransLWPolyLineDirectList(objid1, po1)
            lengthlist, ness = acad.TransNRectLengthListFromSheet(objid1, objid2, po1)
            anglelist = acad.TransLWPolyLineDotAngleList(objid1, po1)
            lengthlist, abcd = nj_to_bc(lengthlist, anglelist, depth)
            pt1 = acad.Vec3Add(ptlist[0], dr0)
            ptlist = [pt1]
            for length, direct in zip(lengthlist, drlist):
                dr1 = acad.Vec3ResetLength(direct, length)
                pt1 = acad.Vec3Add(pt1, dr1)
                ptlist.append(pt1)
            acad.AddLWPolyLine(ptlist)



            
@acad.decorator_command
def llsheet_to_wj_for():
    po1, po2 = acad.GetPoint2()
    if po1 == None: return
    dr0 = acad.Direct(po1, po2)
    while True:
        objidlist = acad.SSGetIdList([[0, "LWPOLYLINE"]])
        if objidlist == None: return
        with acad.transaction() as trans:
            objid1, objid2 = objidlist[0:2]
            objid3, ness = acad.TransAutoFindMidLWPolyLine(objid1, objid2)
            ptlist = acad.TransLWPolyLinePointList(objid3, po1)
            lnlist = acad.TransLWPolyLineLengthList(objid3, po1)
            drlist = acad.TransLWPolyLineDirectList(objid3, po1)
            lnlist = zj_to_wj(lnlist, ness)
            pt1 = acad.Vec3Add(ptlist[0], dr0)
            ptlist = [pt1]
            for length, direct in zip(lnlist, drlist):
                dr1 = acad.Vec3ResetLength(direct, length)
                pt1 = acad.Vec3Add(pt1, dr1)
                ptlist.append(pt1)
            acad.AddLWPolyLine(ptlist)
            acad.TransErase(objid3)

@acad.decorator_command
def llsheet_to_n_rect_for():
    po1, po2 = acad.GetPoint2()
    if po1 == None: return
    dr1 = acad.Direct(po1, po2)
    while True:
        objidlist = acad.SSGetIdList([[0, "LWPOLYLINE"]])
        if objidlist == None: return
        with acad.transaction() as trans:
            buflist = acad.TransNRectPointListFromSheet(objidlist[0], objidlist[1])
            for pt1, pt2, pt3, pt4 in buflist:
                pt1 = acad.Vec3Add(pt1, dr1)
                pt2 = acad.Vec3Add(pt2, dr1)
                pt3 = acad.Vec3Add(pt3, dr1)
                pt4 = acad.Vec3Add(pt4, dr1)
                acad.AddMKPolyLine([pt1, pt2, pt3, pt4])



@acad.decorator_command
def llsheet_from_nj_for(): 
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
def llsheet_from_wj_for(): 
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
        lengthlist = ll_pl.zhu_llpl_zj_to_nj(lengthlist)
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




# @acad.decorator_command
# def llsweep_pl_nj_to_pl_zj_rect_for(): 
#     import ll_pl
#     ll_pl.zhu_uipl_ness()
#     pt1, pt2 = acad.GetPoint2("请选择基点: ", "请选择终点: ")
#     if pt1 == None: return

#     dr0 = acad.Direct(pt1, pt2)
#     objidlist = acad.SSGetIdList([[0, "LWPOLYLINE"]])
#     buflist = []
#     for objid in objidlist:
#         pt0 = acad.GetStartPoint(objid)
#         pt0 = acad.Vec3Add(pt0, dr0)
#         directlist = acad.GetLWPolyLineDirectList(objid)
#         resultlist = ll_pl.zhu_llpl_nj_to_zj_sweep_rect(pt0, directlist)
#         for ptlist in resultlist: 
#             buflist.append(ptlist)

#     with acad.transaction() as trans:
#         for ptlist in buflist:
#             acad.AddLWPolyLine(ptlist) 


# def zhu_llpl_nj_to_zj_sweep_rect(pt0, directlist):
#     count = len(directlist)
#     pt1 = pt0
#     buflist = []
#     for i in range(count):
#         direct = directlist[i]
#         pt2 = acad.Vec3Add(pt1, direct)

#         perdr1, perdr2 = acad.GetPerDirect2ResetLengthXY(pt1, pt2, zhu_llpl_ness_half)
#         po1 = acad.Vec3Add(pt1, perdr1)
#         po2 = acad.Vec3Add(pt2, perdr1)
#         po3 = acad.Vec3Add(pt2, perdr2)
#         po4 = acad.Vec3Add(pt1, perdr2)
#         buflist.append( [po1, po2, po3, po4, po1] )

#         if i < count-1:
#             dr1 = directlist[i]
#             dr2 = directlist[i+1]
#             mo1 = acad.Vec3ResetLength(dr1, zhu_llpl_ness_half)
#             mo2 = acad.Vec3ResetLength(dr2, zhu_llpl_ness_half)
#             dr3 = acad.Vec3Add(mo1, mo2)
#             pt1 = acad.Vec3Add(pt2, dr3)

#     return buflist