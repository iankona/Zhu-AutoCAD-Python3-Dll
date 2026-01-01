import clr

import acad
import academit

import System

def 命令(): 
    academit.添加命令("llpl-to-midpl", llpl_to_midpl)
    academit.添加命令("llpl-print", llpl_print)
    academit.添加命令("llpl-add", llpl_add)
    academit.添加命令("llpl-findstart", llpl_findstart)
    # academit.添加命令("llpl-change-copy-pl", llpl_change_copy_pl)
    academit.添加命令("llpl-sweep", llpl_sweep)
    academit.添加命令("llpl-sweep-set", llpl_sweep_set)
    academit.添加命令("llpl-change-xy-to-xz-plfor", llpl_change_xy_to_xz_plfor)
    academit.添加命令("llpl-change-xy-to-xz-4pl", llpl_change_xy_to_xz_4pl)
    pass


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
            pline_normal = acad.GetLWPolyLineNormal(objid)
            pline_point_list = acad.GetLWPolyLinePointList(objid)
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
            pt0 = acad.GetStartPoint(objid)
            mid = acad.GetLWPolyLineStartMid(objid)
            acad.AddText(pt0, "起点")
            acad.AddText(mid, "中点")


@acad.decorator_command
def llpl_findpoint():
    objidlist = acad.SSGetIdList([[0, "LWPOLYLINE"]])
    with acad.transaction() as trans:
        for objid in objidlist:
            pt0 = acad.GetStartPoint(objid)
            mid = acad.GetLWPolyLineStartMid(objid)
            ptlist = acad.GetLWPolyLinePointList(objid)
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
        mid = acad.GetLWPolyLineStartMid(objid)
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