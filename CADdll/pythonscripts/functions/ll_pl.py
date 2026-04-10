import clr

import acad
import academit

import System

def 命令(): 
    academit.添加命令("llpl-zj-to-nj-y-forward-for", llpl_zj_to_nj_y_forward_for)
    academit.添加命令("llpl-zj-to-nj-y-backward-for", llpl_zj_to_nj_y_backward_for)
    academit.添加命令("llpl-zj-to-zj-y-forward-for", llpl_zj_to_zj_y_forward_for)
    academit.添加命令("llpl-zj-to-zj-y-backward-for", llpl_zj_to_zj_y_backward_for)



    # academit.添加命令("llpl-to-midpl", llpl_to_midpl)
    # academit.添加命令("llpl-print", llpl_print)
    # academit.添加命令("llpl-add", llpl_add)
    # academit.添加命令("llpl-findstart", llpl_findstart)
    # academit.添加命令("llpl-change-copy-pl", llpl_change_copy_pl)
    # academit.添加命令("llpl-sweep", llpl_sweep)
    # academit.添加命令("llpl-sweep-set", llpl_sweep_set)
    # academit.添加命令("llpl-change-xy-to-xz-plfor", llpl_change_xy_to_xz_plfor)
    # academit.添加命令("llpl-change-xy-to-xz-4pl", llpl_change_xy_to_xz_4pl)
    # academit.添加命令("llpl-loft", llpl_loft)
    pass


zhu_llpl_offset = 5000
zhu_llpl_ness = 1.0
zhu_llpl_ness_half = 0.5
zhu_llpl_ness_double = 2.0



def zhu_uipl_offset():
    global zhu_llpl_offset
    offs = acad.GetDouble(zhu_llpl_offset, "请输入起点偏移量:")
    if offs == None: offs = 0
    zhu_llpl_offset = offs


def zhu_uipl_ness():
    global zhu_llpl_ness, zhu_llpl_ness_half, zhu_llpl_ness_double
    ness = acad.GetDouble(zhu_llpl_ness, "请输入板厚:")
    if ness == None: ness = 1.0
    zhu_llpl_ness = ness
    zhu_llpl_ness_half = ness/2
    zhu_llpl_ness_double = ness*2

def zhu_llpl_zj_to_nj(lengthlist):
    count = len(lengthlist)
    match count:
        case 0: return []
        case 1: lengthlist[0] = lengthlist[0] - zhu_llpl_ness
        case _:
            lengthlist[0] = lengthlist[0] - zhu_llpl_ness_half
            for i in range(1, count-1): 
                lengthlist[i] = lengthlist[i] - zhu_llpl_ness
            lengthlist[-1] = lengthlist[-1] - zhu_llpl_ness_half
    return lengthlist

def zhu_llpl_nj_to_zj(lengthlist):
    count = len(lengthlist)
    match count:
        case 0: return []
        case 1: lengthlist[0] = lengthlist[0] + zhu_llpl_ness
        case _:
            lengthlist[0] = lengthlist[0] + zhu_llpl_ness_half
            for i in range(1, count-1): 
                lengthlist[i] = lengthlist[i] + zhu_llpl_ness
            lengthlist[-1] = lengthlist[-1] + zhu_llpl_ness_half
    return lengthlist

def zhu_llpl_wj_to_nj(lengthlist):
    count = len(lengthlist)
    match count:
        case 0: return []
        case 1: lengthlist[0] = lengthlist[0] - zhu_llpl_ness_double
        case _:
            lengthlist[0] = lengthlist[0] - zhu_llpl_ness
            for i in range(1, count-1): 
                lengthlist[i] = lengthlist[i] - zhu_llpl_ness_double
            lengthlist[-1] = lengthlist[-1] - zhu_llpl_ness
    return lengthlist

def zhu_llpl_nj_to_wj(lengthlist):
    count = len(lengthlist)
    match count:
        case 0: return []
        case 1: lengthlist[0] = lengthlist[0] + zhu_llpl_ness_double
        case _:
            lengthlist[0] = lengthlist[0] + zhu_llpl_ness
            for i in range(1, count-1): 
                lengthlist[i] = lengthlist[i] + zhu_llpl_ness_double
            lengthlist[-1] = lengthlist[-1] + zhu_llpl_ness
    return lengthlist

def zhu_llpl_zhanping_y(pt0, lengthlist):
    pt1 = acad.Vec3Add(pt0, [0, zhu_llpl_offset, 0])
    ptlist = [pt1]
    for length in lengthlist:
        pt1 = acad.Vec3Add(pt1, [0, length, 0])
        ptlist.append(pt1)
    return ptlist

def zhu_llpl_zhanping_x(pt0, lengthlist):
    pt1 = acad.Vec3Add(pt0, [zhu_llpl_offset, 0, 0])
    ptlist = [pt1]
    for length in lengthlist:
        pt1 = acad.Vec3Add(pt1, [length, 0, 0])
        ptlist.append(pt1)
    return ptlist



def zhu_llpl_build_pl(pt0, directlist, lengthlist):
    pt1 = pt0
    ptlist = [pt1]
    for direct, length in zip(directlist, lengthlist):
        dr1 = acad.Vec3ResetLength(direct, length)
        pt1 = acad.Vec3Add(pt1, dr1)
        ptlist.append(pt1)
    return ptlist


def zhu_llpl_nj_to_zj_sweep_rect(pt0, directlist):
    count = len(directlist)
    pt1 = pt0
    buflist = []
    for i in range(count):
        direct = directlist[i]
        pt2 = acad.Vec3Add(pt1, direct)

        perdr1, perdr2 = acad.GetPerDirect2ResetLengthXY(pt1, pt2, zhu_llpl_ness_half)
        po1 = acad.Vec3Add(pt1, perdr1)
        po2 = acad.Vec3Add(pt2, perdr1)
        po3 = acad.Vec3Add(pt2, perdr2)
        po4 = acad.Vec3Add(pt1, perdr2)
        buflist.append( [po1, po2, po3, po4, po1] )

        if i < count-1:
            dr1 = directlist[i]
            dr2 = directlist[i+1]
            mo1 = acad.Vec3ResetLength(dr1, zhu_llpl_ness_half)
            mo2 = acad.Vec3ResetLength(dr2, zhu_llpl_ness_half)
            dr3 = acad.Vec3Add(mo1, mo2)
            pt1 = acad.Vec3Add(pt2, dr3)

    return buflist




@acad.decorator_command
def llpl_zj_to_nj_y_forward_for():
    zhu_uipl_offset()
    zhu_uipl_ness()
    objidlist = acad.SSGetIdList([[0, "LWPOLYLINE"]])
    buflist = []
    for objid in objidlist:
        pt0 = acad.GetStartPoint(objid)
        lengthlist = acad.GetLWPolyLineLengthList(objid)
        lengthlist = zhu_llpl_zj_to_nj(lengthlist)
        ptlist = zhu_llpl_zhanping_y(pt0, lengthlist)
        buflist.append(ptlist)

    with acad.transaction() as trans:
        for ptlist in buflist:
            acad.AddLWPolyLine(ptlist)    

@acad.decorator_command
def llpl_zj_to_nj_y_backward_for():
    zhu_uipl_offset()
    zhu_uipl_ness()
    objidlist = acad.SSGetIdList([[0, "LWPOLYLINE"]])
    buflist = []
    for objid in objidlist:
        pt0 = acad.GetStartPoint(objid)
        lengthlist = acad.GetLWPolyLineLengthList(objid)
        lengthlist = zhu_llpl_zj_to_nj(lengthlist) 
        ptlist = zhu_llpl_zhanping_y(pt0, lengthlist[::-1])
        buflist.append(ptlist)

    with acad.transaction() as trans:
        for ptlist in buflist:
            acad.AddLWPolyLine(ptlist)    

@acad.decorator_command
def llpl_zj_to_zj_y_forward_for():
    zhu_uipl_offset()
    zhu_uipl_ness()
    objidlist = acad.SSGetIdList([[0, "LWPOLYLINE"]])
    buflist = []
    for objid in objidlist:
        pt0 = acad.GetStartPoint(objid)
        lengthlist = acad.GetLWPolyLineLengthList(objid)
        ptlist = zhu_llpl_zhanping_y(pt0, lengthlist)
        buflist.append(ptlist)

    with acad.transaction() as trans:
        for ptlist in buflist:
            acad.AddLWPolyLine(ptlist)  

@acad.decorator_command
def llpl_zj_to_zj_y_backward_for():
    zhu_uipl_offset()
    zhu_uipl_ness()
    objidlist = acad.SSGetIdList([[0, "LWPOLYLINE"]])
    buflist = []
    for objid in objidlist:
        pt0 = acad.GetStartPoint(objid)
        lengthlist = acad.GetLWPolyLineLengthList(objid)
        ptlist = zhu_llpl_zhanping_y(pt0, lengthlist[::-1])
        buflist.append(ptlist)

    with acad.transaction() as trans:
        for ptlist in buflist:
            acad.AddLWPolyLine(ptlist)  



@acad.decorator_command
def llpl_to_midpl():
    objidlist = acad.SSGetIdList([[0, "LWPOLYLINE"]])
    with acad.transaction() as trans:
        for objid in objidlist:
            ptlist = []
            pline_point_list = acad.GetLWPolyLinePointList(objid)
            ptlist.append(pline_point_list[0])
            for i in range(len(pline_point_list)-1):
                pt1 = pline_point_list[i]
                pt2 = pline_point_list[i+1]
                ptlist.append(acad.MidPt1Pt2(pt1, pt2))
            ptlist.append(pline_point_list[-1])
            acad.AddLWPolyLine(ptlist)


@acad.decorator_command
def llpl_print():
    objidlist = acad.SSGetIdList([[0, "LWPOLYLINE"]])
    with acad.transaction() as trans:
        for objid in objidlist:
            pline_normal = acad.TransLWPolyLineNormal(objid)
            pline_point_list = acad.TransLWPolyLinePointList(objid)
            print(objid, pline_normal, pline_point_list)



@acad.decorator_command
def llpl_add():
    ptlist = [[17.52426053837262, 56.66072770213942, 68.48634887959271], 
            [17.52426053837261, 56.66072770213942, 13.347482067488855], 
            [17.52426053837261, 73.02319680609772, 13.347482067488855], 
            [17.52426053837262, 73.02319680609772, 68.48634887959271], 
            [17.52426053837262, 56.66072770213942, 68.48634887959271]]
    
    with acad.transaction() as trans:
        pline = acad.AddPolyline3d(ptlist)
        pline.Closed = True

@acad.decorator_command
def llpl_findstart():
    objidlist = acad.SSGetIdList([[0, "LWPOLYLINE"]])
    with acad.transaction() as trans:
        for objid in objidlist:
            pt0 = acad.TransStartPoint(objid)
            mid = acad.TransLWPolyLineStartMid(objid)
            acad.AddText(pt0, "起点")
            acad.AddText(mid, "中点")


@acad.decorator_command
def llpl_findpoint():
    objidlist = acad.SSGetIdList([[0, "LWPOLYLINE"]])
    with acad.transaction() as trans:
        for objid in objidlist:
            pt0 = acad.TransStartPoint(objid)
            mid = acad.TransLWPolyLineStartMid(objid)
            ptlist = acad.TransLWPolyLinePointList(objid)
            for i, pt1 in enumerate(ptlist):
                acad.AddText(pt0, str(i))



@acad.decorator_command
def llpl_change_copy_pl():
    objid = acad.EntSel(string="请点击复制对象: ")
    with acad.transaction() as trans:
        mid = acad.GetLWPolyLineStartMid(objid)
        acad.AddText(mid, "起点")
        pt0 = acad.GetStartPoint(objid)
        drlist = acad.GetLWPolyLineDirectList(objid)
        buflist = [pt0] + drlist
        result1 = acad.ChangeCoordinateXY(buflist, "-Y",  "X")
        result2 = acad.ChangeCoordinateXY(buflist, "-X", "-Y")
        result3 = acad.ChangeCoordinateXY(buflist,  "Y", "-X")
        pt1, dr1list = result1[0], result1[1:]
        pt2, dr2list = result2[0], result2[1:]
        pt3, dr3list = result3[0], result3[1:]
        pt1list = acad.DirectListToPointList(pt1, dr1list)
        pt2list = acad.DirectListToPointList(pt2, dr2list)
        pt3list = acad.DirectListToPointList(pt3, dr3list)
        acad.AddPolyline3d(pt1list)
        acad.AddPolyline3d(pt2list)
        acad.AddPolyline3d(pt3list)
         





@acad.decorator_command
def llpl_change_xy_to_xz_pl_for():
    # objid = acad.EntSel(string="请点击复制对象: ")
    objidlist = acad.SSGetIdList([[0, "LWPOLYLINE"]])
    with acad.transaction() as trans:
        for objid in objidlist:
            ptlist = acad.GetLWPolyLinePointList(objid)
            result = []
            for x,y,z in ptlist:
                result.append([x,z,y])
            acad.AddPolyline3d(result)




@acad.decorator_command
def llpl_change_xy_to_xz_4pl():
    objid = acad.EntSel(string="请点击复制对象: ") # for 循环出错，暂时找不到原因
    with acad.transaction() as trans:
        mid = acad.TransLWPolyLineStartMid(objid)
        acad.AddText(mid, "起点")
        pt0 = acad.GetStartPoint(objid)
        drlist = acad.GetLWPolyLineDirectList(objid)
        drlist = acad.Vec3XYtoXZ(drlist)
        buflist = [pt0] + drlist
        result1 = acad.ChangeCoordinateXY(buflist, "-Y",  "X")
        result2 = acad.ChangeCoordinateXY(buflist, "-X", "-Y")
        result3 = acad.ChangeCoordinateXY(buflist,  "Y", "-X")
        pt0, dr0list = pt0, drlist
        pt1, dr1list = result1[0], result1[1:]
        pt2, dr2list = result2[0], result2[1:]
        pt3, dr3list = result3[0], result3[1:]
        pt0list = acad.DirectListToPointList(pt0, dr0list)
        pt1list = acad.DirectListToPointList(pt1, dr1list)
        pt2list = acad.DirectListToPointList(pt2, dr2list)
        pt3list = acad.DirectListToPointList(pt3, dr3list)
        acad.AddPolyline3d(pt0list)
        acad.AddPolyline3d(pt1list)
        acad.AddPolyline3d(pt2list)
        acad.AddPolyline3d(pt3list)

