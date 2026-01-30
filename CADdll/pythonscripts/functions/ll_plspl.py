import clr

import acad
import academit

import System

def 命令(): 
    academit.添加命令("llpl-zj-to-nj-y-forward-for", llpl_zj_to_nj_y_forward_for)
    academit.添加命令("llpl-zj-to-nj-y-backward-for", llpl_zj_to_nj_y_backward_for)
    academit.添加命令("llpl-zj-to-zj-y-forward-for", llpl_zj_to_zj_y_forward_for)
    academit.添加命令("llpl-zj-to-zj-y-backward-for", llpl_zj_to_zj_y_backward_for)

    academit.添加命令("llpl-connet-pl-and-pl-for", llpl_connet_pl_and_pl_for)

    academit.添加命令("llpl-to-midpl", llpl_to_midpl)
    academit.添加命令("llpl-print", llpl_print)
    academit.添加命令("llpl-add", llpl_add)
    academit.添加命令("llpl-findstart", llpl_findstart)
    # academit.添加命令("llpl-change-copy-pl", llpl_change_copy_pl)
    academit.添加命令("llpl-sweep", llpl_sweep)
    academit.添加命令("llpl-sweep-set", llpl_sweep_set)
    academit.添加命令("llpl-change-xy-to-xz-plfor", llpl_change_xy_to_xz_plfor)
    academit.添加命令("llpl-change-xy-to-xz-4pl", llpl_change_xy_to_xz_4pl)
    academit.添加命令("llpl-loft", llpl_loft)
    pass


zhu_llpl_offset = 5000
zhu_llpl_ness = 1.0
zhu_llpl_ness_half = 0.5
zhu_llpl_ness_double = 2.0
def zhu_uipl_zj_to_nj():
    global zhu_llpl_offset, zhu_llpl_ness, zhu_llpl_ness_half, zhu_llpl_ness_double
    offs = acad.GetDouble(zhu_llpl_offset, "请输入起点偏移量:")
    ness = acad.GetDouble(zhu_llpl_ness, "请输入板厚:")
    if offs == None: offs = 0
    if ness == None: ness = 1.0
    zhu_llpl_offset = offs
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

@acad.decorator_command
def llpl_zj_to_nj_y_forward_for():
    zhu_uipl_zj_to_nj()
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
    zhu_uipl_zj_to_nj()
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
    zhu_uipl_zj_to_nj()
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
    zhu_uipl_zj_to_nj()
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
def llpl_connet_pl_and_pl_for():
    while True:
        objid1 = acad.EntSel("请点击第1条对象:") 
        if acad.IsNone(objid1): break # No method matches given arguments for ObjectId.op_Equality: (<class 'NoneType'>) == 比较出错
        objid2 = acad.EntSel("请点击第2条对象:")
        if acad.IsNone(objid2): break 
        pt1, pd1 = acad.GetStartFinalPoint(objid1)
        lengthlist1 = acad.GetLWPolyLineLengthList(objid1)
        pt2, pd2 = acad.GetStartFinalPoint(objid2)
        lengthlist2 = acad.GetLWPolyLineLengthList(objid2)

        dy1 = acad.Direct(pt1, pd1)
        dy1 = acad.Vec3ResetLength(dy1, 2000)
        pt1 = acad.Vec3Add(pt1, dy1)
        pt2 = acad.Vec3Add(pt2, dy1)
        pd1 = acad.Vec3Add(pd1, dy1)
        pd2 = acad.Vec3Add(pd2, dy1)
        db1 = acad.Direct(pt1, pd1)
        db2 = acad.Direct(pt2, pd2)
        pe1 = acad.Vec3Add(pd1, db1)
        pe2 = acad.Vec3Add(pd2, db2)
        linelist = [[pt1, pe1, "0"], [pt2, pe2, "0"]]

        pt1, pt2 = acad.GetAttachWDirectPt1Pt2(pt1, pt2, 150)
        linelist.append([pt1, pt2, "0"])

        for i, [length1, length2] in enumerate(zip(lengthlist1, lengthlist2)):
            if abs(length1-length2) < 0.0001:
                dr1 = acad.GetPerDirectResetLengthXY(pt1, pt2, pe1, length1)
                po1 = acad.Vec3Add(pt1, dr1)
                po2 = acad.Vec3Add(pt2, dr1)
            else:
                dr1 = acad.GetPerDirectXY(pt1, pt2, pe1)
                dd1 = acad.Vec3ResetLength(dr1, length1)
                dd2 = acad.Vec3ResetLength(dr1, length2)
                po1 = acad.Vec3Add(pt1, dd1)
                po2 = acad.Vec3Add(pt2, dd2)

            linelist.append([pt1, po1, "0"])
            linelist.append([pt2, po2, "0"])
            linelist.append([po1, po2, "图层1"])

            if i == 14:
                dr1 = acad.Vec3ResetLength(db1, 200)
                pt1 = acad.Vec3Add(po1, dr1)
                pt2 = acad.Vec3Add(po2, dr1)
                linelist.append([pt1, pt2, "图层1"])
            else:
                pt1, pt2 = po1, po2

        with acad.transaction() as trans:
            for pt1, pt2, lname in linelist[:-1]:
                acad.AddLine(pt1, pt2, lname)
            pt1, pt2, lname = linelist[-1]
            acad.AddLine(pt1, pt2)

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
    objid = acad.EntSel("请点击复制对象: ")
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
def llpl_change_xy_to_xz_plfor():
    # objid = acad.EntSel("请点击复制对象: ")
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
    objid = acad.EntSel("请点击复制对象: ") # for 循环出错，暂时找不到原因
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
        pt0 = acad.GetStartPoint(plineid)
        mid = acad.GetLWPolyLineStartMid(plineid)
        acad.AddText(pt0, "起点")
        acad.AddText(mid, "中点")
        pt0 = acad.GetStartPoint(rectid)
        mid = acad.GetLWPolyLineStartMid(rectid)
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


@acad.decorator_command
def llpl_loft():
    pass
    # objidlist = acad.SSGetIdList([[0, "LWPOLYLINE"]])
    # with acad.transaction() as trans:
    #     for objid in objidlist:
    #         acad.Copy(objid, [0,0], [1000,1000])


