import clr

import acad
import academit

import System

def 命令(): 
    academit.添加命令("lldk-dis-circle-for", lldk_dis_circle_for)
    # academit.添加命令("lldk-div-circle-for", lldk_div_circle_for)
    academit.添加命令("lldk-rec-dis-circle-for", lldk_rec_dis_circle_for)

zhu_lldk_d = 5
zhu_lldk_r = 2.5
zhu_lldk_offs = 5.5
zhu_lldk_gapa = 20
zhu_lldk_gapb = 20
zhu_lldk_length_sub = 300
def zhu_uidk_dis_circle():
    global zhu_lldk_d, zhu_lldk_r, zhu_lldk_offs, zhu_lldk_gapa, zhu_lldk_gapb, zhu_lldk_length_sub
    d = acad.GetDouble(zhu_lldk_d, "请输入孔直径:")
    offs = acad.GetDouble(zhu_lldk_offs, "请输入偏线距离:")
    gapa = acad.GetDouble(zhu_lldk_gapa, "请输入GapA:")
    gapb = acad.GetDouble(zhu_lldk_gapb, "请输入GapB:")
    lsub = acad.GetDouble(zhu_lldk_length_sub, "请输入分段长度:")
    if d != None: zhu_lldk_d, zhu_lldk_r = d, d/2
    if offs != None: zhu_lldk_offs = offs
    if gapa != None: zhu_lldk_gapa = gapa
    if gapb != None: zhu_lldk_gapb = gapb
    if lsub != None: zhu_lldk_length_sub = lsub


def zhu_lldk_dis_circle(pt1, pt2, pt3):
    dr1 = acad.GetPerDirectXY(pt1, pt2, pt3) 
    dr1 = acad.Vec3ResetLength(dr1, zhu_lldk_offs)
    po1 = acad.Vec3Add(pt1, dr1)
    po2 = acad.Vec3Add(pt2, dr1)
    pt1, pt2 = acad.GetAttachGapAGapBPointList(po1, po2, zhu_lldk_gapa, zhu_lldk_gapb)
    distance = acad.Distance(pt1, pt2)
    count = round(distance/zhu_lldk_length_sub+0.05)
    if count == 1: count = 2 # 段数
    length_sub = distance/count
    dr1 = acad.Direct(pt1, pt2)
    dr1 = acad.Vec3ResetLength(dr1, length_sub)
    result_center_list = []
    result_center_list.append(pt1)
    if count >= 2:
        for j in range(count-1):
            pt1 = acad.Vec3Add(pt1, dr1)
            result_center_list.append(pt1)
    result_center_list.append(pt2)
    return result_center_list



@acad.decorator_command
def lldk_dis_circle_for():
    zhu_uidk_dis_circle()
    while True:
        # objid, pt3 = acad.EntSel("请点击打孔边: "), acad.GetPoint("请点击打孔大概位置: ")
        # if objid == None or pt3 == None: break
        # pt1, pt2 = acad.GetStartFinalPoint(objid)
        pt1, pt2, pt3 = acad.GetPoint3("请选择第1个顶点: ", "请选择第2个顶点: ", "请点击方向顶点: ")
        if pt1 == None: break
        cirlce_center_list = zhu_lldk_dis_circle(pt1, pt2, pt3)
        with acad.transaction() as trans:
            for pt1 in cirlce_center_list:
                acad.AddCircle(pt1, zhu_lldk_r)  


@acad.decorator_command
def lldk_rec_dis_circle_for():
    zhu_uidk_dis_circle()
    objidlist = acad.SSGetIdList([[0, "LWPOLYLINE"]])
    cirlce_center_list = []
    for objid in objidlist:
        center = acad.GetEntityBoundCenterXY(objid)
        pline_point_list = acad.GetLWPolyLinePointList(objid)
        for i in range(len(pline_point_list)-1):
            pt1 = pline_point_list[i]
            pt2 = pline_point_list[i+1]
            pt3 = center
            result_center_list = zhu_lldk_dis_circle(pt1, pt2, pt3)
            cirlce_center_list += result_center_list
    with acad.transaction() as trans:
        for pt1 in cirlce_center_list:
            acad.AddCircle(pt1, zhu_lldk_r)   











        
