import clr

import acad
import academit

import System

def 命令(): 
    academit.添加命令("lldk-circle-subcount-for", lldk_circle_subcount_for)
    academit.添加命令("lldk-circle-sublength-for", lldk_circle_sublength_for)
    academit.添加命令("lldk-rec-circle-subcount-for", lldk_rec_circle_subcount_for)
    academit.添加命令("lldk-rec-circle-sublength-for", lldk_rec_circle_sublength_for)
    academit.添加命令("lldk-copy-subcount-for", lldk_copy_subcount_for)
    academit.添加命令("lldk-copy-sublength-for", lldk_copy_sublength_for)
    academit.添加命令("lldk-jiaoma-for", lldk_jiaoma_for)
    academit.添加命令("lldk-circle-lujing-subcount-for", lldk_circle_lujing_subcount_for)
    academit.添加命令("lldk-circle-lujing-sublength-for", lldk_circle_lujing_sublength_for)

    
llzhu_dk_offset = 5.5
def uizhu_dk_offset():
    global llzhu_dk_offset
    offset = acad.GetDouble(llzhu_dk_offset, "请输入偏线距离:")
    if offset != None: llzhu_dk_offset = offset

llzhu_dk_r = 2.5
llzhu_dk_d = 5
llzhu_dk_gapa = 20
llzhu_dk_gapb = 20
llzhu_dk_cutcount = 3
llzhu_dk_subcount = 2
llzhu_dk_sublength = 300
def uizhu_dk_sublength():
    global llzhu_dk_r, llzhu_dk_d, llzhu_dk_gapa, llzhu_dk_gapb, llzhu_dk_subcount, llzhu_dk_sublength
    d = acad.GetDouble(llzhu_dk_d, "请输入孔直径:")
    gapa = acad.GetDouble(llzhu_dk_gapa, "请输入GapA:")
    gapb = acad.GetDouble(llzhu_dk_gapb, "请输入GapB:")
    length = acad.GetDouble(llzhu_dk_sublength, "请输入间隔长度:")
    if d != None: llzhu_dk_d, llzhu_dk_r = d, d/2
    if gapa != None: llzhu_dk_gapa = gapa
    if gapb != None: llzhu_dk_gapb = gapb
    if length != None: llzhu_dk_sublength = length

# 与均分不同，打孔不涉及圆直径占用的线段计算，打孔计算中圆直径的影响被忽略，只计算圆心位置
def uizhu_dk_subcount():
    global llzhu_dk_r, llzhu_dk_d, llzhu_dk_gapa, llzhu_dk_gapb, llzhu_dk_cutcount, llzhu_dk_subcount, llzhu_dk_sublength
    d = acad.GetDouble(llzhu_dk_d, "请输入孔直径:")
    gapa = acad.GetDouble(llzhu_dk_gapa, "请输入GapA:")
    gapb = acad.GetDouble(llzhu_dk_gapb, "请输入GapB:")
    count = acad.GetInt(llzhu_dk_cutcount, "请输入打孔数量:")
    if d != None: llzhu_dk_d, llzhu_dk_r = d, d/2
    if gapa != None: llzhu_dk_gapa = gapa
    if gapb != None: llzhu_dk_gapb = gapb
    if count != None: 
        llzhu_dk_cutcount = count
        llzhu_dk_subcount = count-1

def uizhu_dk_copy_sublength():
    global llzhu_dk_gapa, llzhu_dk_gapb, llzhu_dk_cutcount, llzhu_dk_subcount, llzhu_dk_sublength
    gapa = acad.GetDouble(llzhu_dk_gapa, "请输入边到中心的长度:")
    length = acad.GetDouble(llzhu_dk_sublength, "请输入间隔长度:")
    if gapa != None: llzhu_dk_gapa, llzhu_dk_gapb = gapa, gapa
    if length != None: llzhu_dk_sublength = length

def uizhu_dk_copy_subcount():
    global llzhu_dk_gapa, llzhu_dk_gapb, llzhu_dk_cutcount, llzhu_dk_subcount, llzhu_dk_sublength
    gapa = acad.GetDouble(llzhu_dk_gapa, "请输入边到中心的长度:")
    count = acad.GetInt(llzhu_dk_cutcount, "请输入打孔数量:")
    if gapa != None: llzhu_dk_gapa, llzhu_dk_gapb = gapa, gapa
    if count != None: 
        llzhu_dk_cutcount = count
        llzhu_dk_subcount = count-1

def __calc_circle_position(pt1, pt2, pt3, count, length):
    dr1 = acad.GetPerDirectXY(pt1, pt2, pt3) 
    dr1 = acad.Vec3ResetLength(dr1, llzhu_dk_offset)
    po1 = acad.Vec3Add(pt1, dr1)
    po2 = acad.Vec3Add(pt2, dr1)
    lengthlist = [llzhu_dk_gapa] + [length for i in range(count)]
    dr1 = acad.Direct(po1, po2)
    result_center_list = []
    for length in lengthlist:
        dr1 = acad.Vec3ResetLength(dr1, length)
        po1 = acad.Vec3Add(po1, dr1)
        result_center_list.append(po1)
    return result_center_list

@acad.decorator_command
def lldk_circle_sublength_for():
    uizhu_dk_offset()
    uizhu_dk_sublength()
    while True:
        # objid, pt3 = acad.EntSel("请点击打孔边: "), acad.GetPoint("请点击打孔大概位置: ")
        # if objid == None or pt3 == None: break
        # pt1, pt2 = acad.GetStartFinalPoint(objid)
        pt1, pt2, pt3 = acad.GetPoint3("请选择第1个顶点: ", "请选择第2个顶点: ", "请点击方向顶点: ")
        if pt1 == None: break
        distance = acad.Distance(pt1, pt2)
        distance = distance-llzhu_dk_gapa-llzhu_dk_gapb
        count = round(distance/llzhu_dk_sublength) 
        if count < 1: count = 1
        length = distance / count
        cirlce_center_list = __calc_circle_position(pt1, pt2, pt3, count, length)
        with acad.transaction() as trans:
            for pt1 in cirlce_center_list:
                acad.AddCircle(pt1, llzhu_dk_r)  

@acad.decorator_command
def lldk_circle_subcount_for():
    uizhu_dk_offset()
    uizhu_dk_subcount()
    while True:
        # objid, pt3 = acad.EntSel("请点击打孔边: "), acad.GetPoint("请点击打孔大概位置: ")
        # if objid == None or pt3 == None: break
        # pt1, pt2 = acad.GetStartFinalPoint(objid)
        pt1, pt2, pt3 = acad.GetPoint3("请选择第1个顶点: ", "请选择第2个顶点: ", "请点击方向顶点: ")
        if pt1 == None: break
        distance = acad.Distance(pt1, pt2)
        distance = distance-llzhu_dk_gapa-llzhu_dk_gapb
        count = llzhu_dk_subcount
        if count < 1: count = 1
        length = distance / count
        cirlce_center_list = __calc_circle_position(pt1, pt2, pt3, count, length)
        with acad.transaction() as trans:
            for pt1 in cirlce_center_list:
                acad.AddCircle(pt1, llzhu_dk_r)  



@acad.decorator_command
def lldk_rec_circle_sublength_for():
    uizhu_dk_offset()
    uizhu_dk_sublength()
    objidlist = acad.SSGetIdList([[0, "LWPOLYLINE"]])
    cirlce_center_list = []
    for objid in objidlist:
        center = acad.GetEntityBoundCenterXY0(objid)
        pline_point_list = acad.GetLWPolyLinePointList(objid)
        for i in range(len(pline_point_list)-1):
            pt1 = pline_point_list[i]
            pt2 = pline_point_list[i+1]
            pt3 = center
            distance = acad.Distance(pt1, pt2)
            distance = distance-llzhu_dk_gapa-llzhu_dk_gapb
            count = round(distance/llzhu_dk_sublength)
            if count < 1: count = 1
            length = distance / count
            result_center_list = __calc_circle_position(pt1, pt2, pt3, count, length)
            cirlce_center_list += result_center_list
    with acad.transaction() as trans:
        for pt1 in cirlce_center_list:
            acad.AddCircle(pt1, llzhu_dk_r)   


@acad.decorator_command
def lldk_rec_circle_subcount_for():
    uizhu_dk_offset()
    uizhu_dk_sublength()
    objidlist = acad.SSGetIdList([[0, "LWPOLYLINE"]])
    cirlce_center_list = []
    for objid in objidlist:
        center = acad.GetEntityBoundCenterXY0(objid)
        pline_point_list = acad.GetLWPolyLinePointList(objid)
        for i in range(len(pline_point_list)-1):
            pt1 = pline_point_list[i]
            pt2 = pline_point_list[i+1]
            pt3 = center
            distance = acad.Distance(pt1, pt2)
            distance = distance-llzhu_dk_gapa-llzhu_dk_gapb
            count = llzhu_dk_subcount
            if count < 1: count = 1
            length = distance / count
            result_center_list = __calc_circle_position(pt1, pt2, pt3, count, length)
            cirlce_center_list += result_center_list
    with acad.transaction() as trans:
        for pt1 in cirlce_center_list:
            acad.AddCircle(pt1, llzhu_dk_r)   





@acad.decorator_command
def lldk_copy_sublength_for():
    uizhu_dk_offset()
    uizhu_dk_copy_sublength()
    pt1, pt2 = acad.GetPoint2("请选择打孔图像第1点:", "请选择打孔图像第2点:")
    mid = acad.MidPt1Pt2(pt1, pt2)
    objid = acad.GetSelectCornerCrossIdList(pt1, pt2)[0]
    while True:
        dr1 = acad.Direct(pt1, pt2)
        po1, po2, po3 = acad.GetPoint3()
        dr2 = acad.Direct(po1, po2)
        angle = acad.AngleFromDotDr1Dr2(dr1, dr2)
        axis = acad.CrossNormalized(dr1, dr2)
        distance = acad.Distance(po1, po2)
        distance = distance-llzhu_dk_gapa-llzhu_dk_gapb
        count = round(distance/llzhu_dk_sublength)
        length = distance / count
        perdr1 = acad.GetPerDirectResetLengthXY(po1, po2, po3, llzhu_dk_offset)
        dr2 = acad.Vec3ResetLength(dr2, llzhu_dk_gapa)
        po1 = acad.Vec3Add(po1, perdr1)
        po1 = acad.Vec3Add(po1, dr2)
        dr2 = acad.Vec3ResetLength(dr2, length)
        with acad.transaction() as trans:
            for i in range(count+1):
                copy = acad.TransRoationCopy(objid, angle, axis, mid)
                acad.TransMove(copy.ObjectId, mid, po1)
                po1 = acad.Vec3Add(po1, dr2)



@acad.decorator_command
def lldk_copy_subcount_for():
    uizhu_dk_offset()
    uizhu_dk_copy_subcount()
    pt1, pt2 = acad.GetPoint2("请选择打孔图像第1点:", "请选择打孔图像第2点:")
    mid = acad.MidPt1Pt2(pt1, pt2)
    objid = acad.GetSelectCornerCrossIdList(pt1, pt2)[0]
    while True:
        dr1 = acad.Direct(pt1, pt2)
        po1, po2, po3 = acad.GetPoint3()
        dr2 = acad.Direct(po1, po2)
        angle = acad.AngleFromDotDr1Dr2(dr1, dr2)
        axis = acad.CrossNormalized(dr1, dr2)
        distance = acad.Distance(po1, po2)
        distance = distance-llzhu_dk_gapa-llzhu_dk_gapb
        count = llzhu_dk_subcount
        if count < 1: count = 1
        length = distance / count
        perdr1 = acad.GetPerDirectResetLengthXY(po1, po2, po3, llzhu_dk_offset)
        dr2 = acad.Vec3ResetLength(dr2, llzhu_dk_gapa)
        po1 = acad.Vec3Add(po1, perdr1)
        po1 = acad.Vec3Add(po1, dr2)
        dr2 = acad.Vec3ResetLength(dr2, length)
        with acad.transaction() as trans:
            for i in range(count+1):
                copy = acad.TransRoationCopy(objid, angle, axis, mid)
                acad.TransMove(copy.ObjectId, mid, po1)
                po1 = acad.Vec3Add(po1, dr2)




@acad.decorator_command
def lldk_jiaoma_for():
    uizhu_dk_offset()
    uizhu_dk_copy_sublength()
    pt1, pt2 = acad.GetPoint2("请选择打孔图像第1点:", "请选择打孔图像第2点:")
    mid = acad.MidPt1Pt2(pt1, pt2)
    objidlist = acad.GetSelectCornerCrossIdList(pt1, pt2)
    while True:
        dr1 = acad.Direct(pt1, pt2)
        po1, po2, po3 = acad.GetPoint3()
        dr2 = acad.Direct(po1, po2)
        angle = acad.AngleFromDotDr1Dr2(dr1, dr2)
        axis = acad.CrossNormalized(dr1, dr2)
        distance = acad.Distance(po1, po2)
        distance = distance-llzhu_dk_gapa
        count = int(distance/llzhu_dk_sublength)
        perdr1 = acad.GetPerDirectResetLengthXY(po1, po2, po3, llzhu_dk_offset)
        dr2 = acad.Vec3ResetLength(dr2, llzhu_dk_gapa)
        po1 = acad.Vec3Add(po1, perdr1)
        po1 = acad.Vec3Add(po1, dr2)
        dr2 = acad.Vec3ResetLength(dr2, llzhu_dk_sublength)
        with acad.transaction() as trans:
            for i in range(count+1):
                copyidlist = acad.TransRoationCopyIdList(objidlist, angle, axis, mid)
                acad.TransMoveIdList(copyidlist, mid, po1)
                po1 = acad.Vec3Add(po1, dr2)









def uizhu_dk_lujing_sublength():
    global llzhu_dk_r, llzhu_dk_d, llzhu_dk_gapa, llzhu_dk_gapb, llzhu_dk_subcount, llzhu_dk_sublength
    d = acad.GetDouble(llzhu_dk_d, "请输入孔直径:")
    gapa = acad.GetDouble(llzhu_dk_gapa, "请输入起点到中心的长度:")
    length = acad.GetDouble(llzhu_dk_sublength, "请输入间隔长度:")
    if d != None: llzhu_dk_d, llzhu_dk_r = d, d/2
    if gapa != None: llzhu_dk_gapa, llzhu_dk_gapb = gapa, gapa
    if length != None: llzhu_dk_sublength = length

# 与均分不同，打孔不涉及圆直径占用的线段计算，打孔计算中圆直径的影响被忽略，只计算圆心位置
def uizhu_dk_lujing_subcount():
    global llzhu_dk_r, llzhu_dk_d, llzhu_dk_gapa, llzhu_dk_gapb, llzhu_dk_cutcount, llzhu_dk_subcount, llzhu_dk_sublength
    d = acad.GetDouble(llzhu_dk_d, "请输入孔直径:")
    gapa = acad.GetDouble(llzhu_dk_gapa, "请输入起点到中心的长度:")
    count = acad.GetInt(llzhu_dk_cutcount, "请输入打孔数量:")
    if d != None: llzhu_dk_d, llzhu_dk_r = d, d/2
    if gapa != None: llzhu_dk_gapa, llzhu_dk_gapb = gapa, gapa
    if count != None: 
        llzhu_dk_cutcount = count
        llzhu_dk_subcount = count-1

        
@acad.decorator_command
def lldk_circle_lujing_subcount_for():
    uizhu_dk_lujing_subcount()
    objidlist = acad.SSGetIdList([[0, "LWPOLYLINE"]])
    with acad.transaction() as trans:
        for objid in objidlist: 
            objref = acad.TransObjectForRead(objid)
            distance = objref.Length
            subcount = llzhu_dk_subcount
            if subcount < 2: subcount = 2
            sublength = distance/subcount
            gapa = llzhu_dk_gapa
            if gapa > sublength: gapa = sublength
            pointlist = []
            sumlength = gapa
            for i in range(subcount):
                point = objref.GetPointAtDist(sumlength)
                pointlist.append(point)
                sumlength += sublength

            ptlist = []
            for point in pointlist: ptlist.append([point.X, point.Y, point.Z])

    with acad.transaction() as trans:
        for pt1 in ptlist:
            acad.AddCircle(pt1, llzhu_dk_r)  

@acad.decorator_command
def lldk_circle_lujing_sublength_for():
    uizhu_dk_lujing_sublength()
    objidlist = acad.SSGetIdList([[0, "LWPOLYLINE"]])
    with acad.transaction() as trans:
        for objid in objidlist: 
            objref = acad.TransObjectForRead(objid)
            distance = objref.Length
            subcount = round(distance/llzhu_dk_sublength)
            if subcount < 2: subcount = 2
            sublength = distance/subcount
            gapa = llzhu_dk_gapa
            if gapa > sublength: gapa = sublength
            pointlist = []
            sumlength = gapa
            for i in range(subcount):
                point = objref.GetPointAtDist(sumlength)
                pointlist.append(point)
                sumlength += sublength

            ptlist = []
            for point in pointlist: ptlist.append([point.X, point.Y, point.Z])

    with acad.transaction() as trans:
        for pt1 in ptlist:
            acad.AddCircle(pt1, llzhu_dk_r)  