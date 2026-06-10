import clr

import acad
import academit

import System

def 命令(): 
    academit.添加命令("llyx3d-for", llyx3d_for)
    academit.添加命令("llyx3d-side3-for", llyx3d_side3_for)
    academit.添加命令("llyx3d-region-side3-for", llyx3d_region_side3_for)
    academit.添加命令("llyx3d-region-from3-for", llyx3d_region_from3_for)
    academit.添加命令("llyx3d-region-from4-for", llyx3d_region_from4_for)
    academit.添加命令("llyx3d-region-from6-for", llyx3d_region_from6_for)
    academit.添加命令("llyx3d-region-fromx-for", llyx3d_region_fromx_for)



llzhu_yx_offset = 50
def zhu_ui_llyx_offset():
    global llzhu_yx_offset
    offset = acad.GetDouble(llzhu_yx_offset, "请输入偏移大小: ")
    if offset != None: llzhu_yx_offset = offset


@acad.decorator_command
def llyx3d_for():
    zhu_ui_llyx_offset()
    while True:
        pt1, pt2, pt3 = acad.GetPoint3("请选择第1个顶点: ", "请选择第2个顶点: ", "请点击方向顶点: ")
        if pt1 == None: break
        with acad.transaction() as trans:
            line1 = acad.Line(acad.ToPoint3d(pt1), acad.ToPoint3d(pt2))
            point = line1.GetClosestPointTo(acad.ToPoint3d(pt3), extend=True)
            po1 = [point.X, point.Y, point.Z]
            dr1 = acad.Direct(po1, pt3)
            dr1 = acad.Vec3ResetLength(dr1, llzhu_yx_offset)
            po1 = acad.Vec3Add(pt1, dr1)
            po2 = acad.Vec3Add(pt2, dr1)
            acad.AddLine(po1, po2)

@acad.decorator_command
def llyx3d_side3_for():
    zhu_ui_llyx_offset()
    while True:
        pt1, pt2, pt3 = acad.GetPoint3("请选择第1个顶点: ", "请选择第2个顶点: ", "请点击方向顶点: ")
        if pt1 == None: break
        with acad.transaction() as trans:
            line1 = acad.Line(acad.ToPoint3d(pt1), acad.ToPoint3d(pt2))
            point = line1.GetClosestPointTo(acad.ToPoint3d(pt3), extend=True)
            po1 = [point.X, point.Y, point.Z]
            dr1 = acad.Direct(po1, pt3)
            dr1 = acad.Vec3ResetLength(dr1, llzhu_yx_offset)
            po1 = acad.Vec3Add(pt1, dr1)
            po2 = acad.Vec3Add(pt2, dr1)
            acad.AddLine(pt1, po1)
            acad.AddLine(pt2, po2)
            acad.AddLine(po1, po2)





@acad.decorator_command
def llyx3d_region_side3_for():
    zhu_ui_llyx_offset()
    while True:
        pt1, pt2, pt3 = acad.GetPoint3("请选择第1个顶点: ", "请选择第2个顶点: ", "请点击方向顶点: ")
        if pt1 == None: break
        with acad.transaction() as trans:
            line1 = acad.DBObjectLine(pt1, pt2)
            point = line1.GetClosestPointTo(acad.ToPoint3d(pt3), extend=True)
            po1 = [point.X, point.Y, point.Z]
            dr1 = acad.Direct(po1, pt3)
            dr1 = acad.Vec3ResetLength(dr1, llzhu_yx_offset)
            po1 = acad.Vec3Add(pt1, dr1)
            po2 = acad.Vec3Add(pt2, dr1)
            line3 = acad.DBObjectLine(pt2, po2)
            line4 = acad.DBObjectLine(po2, po1)
            line2 = acad.DBObjectLine(po1, pt1)
            buflist = [line1, line2, line3, line4]
            collect = acad.DBObjectCollection()
            for line1 in buflist: collect.Add(line1)
            regions = acad.Region.CreateFromCurves(collect)
            for region in regions: acad.AddDBObject(region)


@acad.decorator_command
def llyx3d_region_from3_for():
    while True:
        pt1, pt2, pt3 = acad.GetPoint3()
        if pt1 == None: break
        with acad.transaction() as trans:
            line1 = acad.DBObjectLine(pt1, pt2)
            line2 = acad.DBObjectLine(pt2, pt3)
            line3 = acad.DBObjectLine(pt3, pt1)
            buflist = [line1, line2, line3]
            collect = acad.DBObjectCollection()
            for line1 in buflist: collect.Add(line1)
            regions = acad.Region.CreateFromCurves(collect)
            for region in regions: acad.AddDBObject(region)


@acad.decorator_command
def llyx3d_region_from4_for():
    while True:
        pt1, pt2, pt3, pt4 = acad.GetPoint4()
        if pt1 == None: break
        with acad.transaction() as trans:
            line1 = acad.DBObjectLine(pt1, pt2)
            line2 = acad.DBObjectLine(pt2, pt3)
            line3 = acad.DBObjectLine(pt3, pt4)
            line4 = acad.DBObjectLine(pt4, pt1)
            buflist = [line1, line2, line3, line4]
            collect = acad.DBObjectCollection()
            for line1 in buflist: collect.Add(line1)
            regions = acad.Region.CreateFromCurves(collect)
            for region in regions: acad.AddDBObject(region)

@acad.decorator_command
def llyx3d_region_from6_for():
    while True:
        pt1, pt2, pt3, pt4 = acad.GetPoint4()
        if pt1 == None: return
        pt5, pt6 = acad.GetPoint2()
        if pt1 == None: return
        with acad.transaction() as trans:
            line1 = acad.DBObjectLine(pt1, pt2)
            line2 = acad.DBObjectLine(pt2, pt3)
            line3 = acad.DBObjectLine(pt3, pt4)
            line4 = acad.DBObjectLine(pt4, pt5)
            line5 = acad.DBObjectLine(pt5, pt6)
            line6 = acad.DBObjectLine(pt6, pt1)
            buflist = [line1, line2, line3, line4, line5, line6]
            collect = acad.DBObjectCollection()
            for line1 in buflist: collect.Add(line1)
            regions = acad.Region.CreateFromCurves(collect)
            for region in regions: acad.AddDBObject(region)



@acad.decorator_command
def llyx3d_region_fromx_for():
    while True:
        ptlist = []
        while True:
            pt1 = acad.GetPoint()
            if pt1 == None: break 
            ptlist.append(pt1)
        if ptlist == []: return
        with acad.transaction() as trans:
            buflist = []
            count = len(ptlist)
            for i in range(count):
                k = i % count
                m = (i+1) % count
                pt1 = ptlist[k]
                pt2 = ptlist[m]
                line1 = acad.DBObjectLine(pt1, pt2)
                buflist.append(line1)

            collect = acad.DBObjectCollection()
            for line1 in buflist: collect.Add(line1)
            regions = acad.Region.CreateFromCurves(collect)
            for region in regions: acad.AddDBObject(region)