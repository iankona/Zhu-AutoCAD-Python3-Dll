import clr

import acad
import academit

import System

def 命令(): 
    academit.添加命令("llyx-for", llyx_for)
    academit.添加命令("llyx-side3-for", llyx_side3_for)
    academit.添加命令("llyx-sidex-for", llyx_sidex_for)
    academit.添加命令("llyx-sidex-n-for", llyx_sidex_n_for)
    academit.添加命令("llyx-sidex-w-for", llyx_sidex_w_for)

    academit.添加命令("llyx-fence-side3-for", llyx_fence_side3_for)
    academit.添加命令("llyx-fence-sidex-for", llyx_fence_sidex_for)
    academit.添加命令("llyx-fence-sidex-n-for", llyx_fence_sidex_n_for)
    academit.添加命令("llyx-fence-sidex-w-for", llyx_fence_sidex_w_for)

    academit.添加命令("llyx-rect-side3-for", llyx_rect_side3_for)
    academit.添加命令("llyx-rect-sidex-for", llyx_rect_sidex_for)
    academit.添加命令("llyx-rect-sidex-n-for", llyx_rect_sidex_n_for)
    academit.添加命令("llyx-rect-sidex-w-for", llyx_rect_sidex_w_for)

    academit.添加命令("llyx-rect-sidex-panjiayuan-for", llyx_rect_sidex_panjiayuan_for)






zhu_llyx_ness_length = 0
def zhu_ui_llyx_ness():
    global zhu_llyx_ness_length
    ness = acad.GetDouble(zhu_llyx_ness_length, "请输入板厚: ")
    if ness != None: zhu_llyx_ness_length = ness

zhu_llyx_offset_length = 50
def zhu_ui_llyx_offset():
    global zhu_llyx_offset_length
    length = acad.GetDouble(zhu_llyx_offset_length, "请输入偏移大小: ")
    if length != None: zhu_llyx_offset_length = length


def zhu_trans_llyx_att_n_line2(pt1, pt2, length): # per direct
    po1, po2 = acad.GetAttachNDirectPointList(pt1, pt2, length)
    line1 = acad.AddLine(pt1, po1)
    line2 = acad.AddLine(pt2, po2)
    return po1, po2, None

def zhu_trans_llyx_att_w_line2(pt1, pt2, length): # per direct
    po1, po2 = acad.GetAttachWDirectPointList(pt1, pt2, length)
    line1 = acad.AddLine(pt1, po1)
    line2 = acad.AddLine(pt2, po2)
    return po1, po2, None

def zhu_trans_llyx_per_p_line3(pt1, pt2, dr1, length): # per direct
    dr1 = acad.Vec3ResetLength(dr1, length)
    po1 = acad.Vec3Add(pt1, dr1)
    po2 = acad.Vec3Add(pt2, dr1)
    line1 = acad.AddLine(pt1, po1)
    line2 = acad.AddLine(pt2, po2)
    line3 = acad.AddLine(po1, po2, "图层1")
    return po1, po2, line3

def zhu_trans_llyx_per_n_line3(pt1, pt2, dr1, length):
    dr1 = acad.Vec3ResetLength(dr1, length)
    po1 = acad.Vec3Add(pt1, dr1)
    po2 = acad.Vec3Add(pt2, dr1)
    po1, po2 = acad.GetAttachNDirectPointList(po1, po2, length)
    line1 = acad.AddLine(pt1, po1)
    line2 = acad.AddLine(pt2, po2)
    if zhu_llyx_ness_length != 0:
        pn1, pn2 = acad.GetAttachNDirectPointList(po1, po2, zhu_llyx_ness_length)
        line1 = acad.AddLine(po1, pn1)
        line2 = acad.AddLine(po2, pn2) 
        po1, po2 = pn1, pn2
    line3 = acad.AddLine(po1, po2, "图层1")
    return po1, po2, line3

def zhu_trans_llyx_per_w_line3(pt1, pt2, dr1, length):
    if zhu_llyx_ness_length != 0:
        po1, po2 = acad.GetAttachWDirectPointList(pt1, pt2, zhu_llyx_ness_length)
        line1 = acad.AddLine(pt1, po1)
        line2 = acad.AddLine(pt2, po2)  
        pt1, pt2 = po1, po2
    dr1 = acad.Vec3ResetLength(dr1, length)
    po1 = acad.Vec3Add(pt1, dr1)
    po2 = acad.Vec3Add(pt2, dr1)
    po1, po2 = acad.GetAttachWDirectPointList(po1, po2, length)
    line1 = acad.AddLine(pt1, po1)
    line2 = acad.AddLine(pt2, po2)
    line3 = acad.AddLine(po1, po2, "图层1")
    return po1, po2, line3

def zhu_trans_llyx_per_p_lengthlist(pt1, pt2, dr1, lengthlist):
    for length in lengthlist:
        pt1, pt2, line3 = zhu_trans_llyx_per_p_line3(pt1, pt2, dr1, length)
    line3.Layer = "0"

def zhu_trans_llyx_per_n_lengthlist(pt1, pt2, dr1, lengthlist):
    for m, length in enumerate(lengthlist):
        if m%2 == 1: 
            pt1, pt2, line3 = zhu_trans_llyx_per_n_line3(pt1, pt2, dr1, length)
        else:
            pt1, pt2, line3 = zhu_trans_llyx_per_p_line3(pt1, pt2, dr1, length)
    line3.Layer = "0"

def zhu_trans_llyx_per_w_lengthlist(pt1, pt2, dr1, lengthlist):
    for m, length in enumerate(lengthlist):
        if m%2 == 1: 
            pt1, pt2, line3 = zhu_trans_llyx_per_w_line3(pt1, pt2, dr1, length)
        else:
            pt1, pt2, line3 = zhu_trans_llyx_per_p_line3(pt1, pt2, dr1, length)
    line3.Layer = "0"



@acad.decorator_command
def llyx_for():
    zhu_ui_llyx_offset()
    while True:
        pt1, pt2, pt3 = acad.GetPoint3("请选择第1个顶点: ", "请选择第2个顶点: ", "请点击方向顶点: ")
        if pt1 == None: break
        with acad.transaction() as trans:
            dr1 = acad.GetPerDirectResetLengthXY(pt1, pt2, pt3, zhu_llyx_offset_length) # 点在线上，WhichSideOfLineXY 会出现cannot access local variable 'result' where it is not associated with a value
            po1 = acad.Vec3Add(pt1, dr1)
            po2 = acad.Vec3Add(pt2, dr1)
            acad.AddLine(po1, po2)

@acad.decorator_command_undo
def llyx_side3_for():
    zhu_ui_llyx_offset()
    while True:
        pt1, pt2, pt3 = acad.GetPoint3("请选择第1个顶点: ", "请选择第2个顶点: ", "请点击方向顶点: ")
        if pt1 == None: break
        with acad.transaction() as trans:
            dr1 = acad.GetPerDirectXY(pt1, pt2, pt3)
            pt1, pt2, line3 = zhu_trans_llyx_per_p_line3(pt1, pt2, dr1, zhu_llyx_offset_length)
            line3.Layer = "0"



@acad.decorator_command
def llyx_sidex_for():
    列表 = acad.GetDoubleListLimitCount()
    if 列表 == None: return
    while True:
        pt1, pt2, pt3 = acad.GetPoint3("请选择第1个顶点: ", "请选择第2个顶点: ", "请点击方向顶点: ")
        if pt1 == None: break
        with acad.transaction() as trans:
            dr1 = acad.GetPerDirectXY(pt1, pt2, pt3)
            zhu_trans_llyx_per_p_lengthlist(pt1, pt2, dr1, 列表)



@acad.decorator_command
def llyx_sidex_n_for():
    zhu_ui_llyx_ness()
    列表 = acad.GetDoubleListLimitCount()
    if 列表 == None: return
    while True:
        pt1, pt2, pt3 = acad.GetPoint3("请选择第1个顶点: ", "请选择第2个顶点: ", "请点击方向顶点: ")
        if pt1 == None: break
        with acad.transaction() as trans:
            dr1 = acad.GetPerDirectXY(pt1, pt2, pt3)
            zhu_trans_llyx_per_n_lengthlist(pt1, pt2, dr1, 列表)




@acad.decorator_command
def llyx_sidex_w_for():
    zhu_ui_llyx_ness()
    列表 = acad.GetDoubleListLimitCount()
    if 列表 == None: return
    while True:
        pt1, pt2, pt3 = acad.GetPoint3("请选择第1个顶点: ", "请选择第2个顶点: ", "请点击方向顶点: ")
        if pt1 == None: break
        with acad.transaction() as trans:
            dr1 = acad.GetPerDirectXY(pt1, pt2, pt3)
            zhu_trans_llyx_per_w_lengthlist(pt1, pt2, dr1, 列表)


@acad.decorator_command
def llyx_rect_side3_for():
    zhu_ui_llyx_offset()
    objidlist = acad.SSGetIdList([[0, "LWPOLYLINE"]])
    with acad.transaction() as trans:
        acad.ChangeObjectIdLayer(objidlist, "图层1")
        result = acad.TransAutoFindRectPointList(objidlist)
        for pt1, pt2, dr1 in result:
            dr1 = acad.Vec3ResetLength(dr1, zhu_llyx_offset_length)
            po1 = acad.Vec3Add(pt1, dr1)
            po2 = acad.Vec3Add(pt2, dr1)
            acad.AddLine(pt1, po1)
            acad.AddLine(pt2, po2)
            acad.AddLine(po1, po2)


@acad.decorator_command
def llyx_rect_sidex_for():
    列表 = acad.GetDoubleListLimitCount()
    if 列表 == None: return
    objidlist = acad.SSGetIdList([[0, "LWPOLYLINE"]])
    with acad.transaction() as trans:
        acad.ChangeObjectIdLayer(objidlist, "图层1")
        result = acad.TransAutoFindRectPointList(objidlist)
        for pt1, pt2, dr1 in result:
            zhu_trans_llyx_per_p_lengthlist(pt1, pt2, dr1, 列表)


@acad.decorator_command
def llyx_rect_sidex_n_for():
    zhu_ui_llyx_ness()
    列表 = acad.GetDoubleListLimitCount()
    if 列表 == []: return
    objidlist = acad.SSGetIdList([[0, "LWPOLYLINE"]])
    with acad.transaction() as trans:
        acad.ChangeObjectIdLayer(objidlist, "图层1")
        result = acad.TransAutoFindRectPointList(objidlist)
        for pt1, pt2, dr1 in result:
            zhu_trans_llyx_per_n_lengthlist(pt1, pt2, dr1, 列表)


@acad.decorator_command
def llyx_rect_sidex_w_for():
    zhu_ui_llyx_ness()
    列表 = acad.GetDoubleListLimitCount()
    if 列表 == None: return
    objidlist = acad.SSGetIdList([[0, "LWPOLYLINE"]])
    with acad.transaction() as trans:
        acad.ChangeObjectIdLayer(objidlist, "图层1")
        result = acad.TransAutoFindRectPointList(objidlist)
        for pt1, pt2, dr1 in result:
            zhu_trans_llyx_per_w_lengthlist(pt1, pt2, dr1, 列表)



@acad.decorator_command
def llyx_fence_side3_for():
    zhu_ui_llyx_offset()
    while True:
        pt1, pt2 = acad.GetPoint2()
        if pt1 == None: return
        objidlist = acad.GetSelectFenceIdList(pt1, pt2, [[0, "LWPOLYLINE"]])
        with acad.transaction() as trans:
            acad.ChangeObjectIdLayer(objidlist, "图层1")
            result = acad.TransAutoFindRectFencePointList(pt1, pt2, objidlist)
            for pt1, pt2, dr1 in result:
                dr1 = acad.Vec3ResetLength(dr1, zhu_llyx_offset_length)
                po1 = acad.Vec3Add(pt1, dr1)
                po2 = acad.Vec3Add(pt2, dr1)
                acad.AddLine(pt1, po1)
                acad.AddLine(pt2, po2)
                acad.AddLine(po1, po2)



@acad.decorator_command
def llyx_fence_sidex_for():
    列表 = acad.GetDoubleListLimitCount()
    if 列表 == None: return
    while True:
        pt1, pt2 = acad.GetPoint2()
        if pt1 == None: return
        objidlist = acad.GetSelectFenceIdList(pt1, pt2, [[0, "LWPOLYLINE"]])
        with acad.transaction() as trans:
            acad.ChangeObjectIdLayer(objidlist, "图层1")
            result = acad.TransAutoFindRectFencePointList(pt1, pt2, objidlist)
            for pt1, pt2, dr1 in result: 
                zhu_trans_llyx_per_p_lengthlist(pt1, pt2, dr1, 列表)


@acad.decorator_command
def llyx_fence_sidex_n_for():
    zhu_ui_llyx_ness()
    列表 = acad.GetDoubleListLimitCount()
    if 列表 == None: return
    while True:
        pt1, pt2 = acad.GetPoint2()
        if pt1 == None: return
        objidlist = acad.GetSelectFenceIdList(pt1, pt2, [[0, "LWPOLYLINE"]])
        with acad.transaction() as trans:
            acad.ChangeObjectIdLayer(objidlist, "图层1")
            result = acad.TransAutoFindRectFencePointList(pt1, pt2, objidlist)
            for pt1, pt2, dr1 in result:
                zhu_trans_llyx_per_n_lengthlist(pt1, pt2, dr1, 列表)




@acad.decorator_command
def llyx_fence_sidex_w_for():
    zhu_ui_llyx_ness()
    列表 = acad.GetDoubleListLimitCount()
    if 列表 == None: return
    while True:
        pt1, pt2 = acad.GetPoint2()
        if pt1 == None: return
        objidlist = acad.GetSelectFenceIdList(pt1, pt2, [[0, "LWPOLYLINE"]])
        with acad.transaction() as trans:
            acad.ChangeObjectIdLayer(objidlist, "图层1")
            result = acad.TransAutoFindRectFencePointList(pt1, pt2, objidlist)
            for pt1, pt2, dr1 in result:
                zhu_trans_llyx_per_w_lengthlist(pt1, pt2, dr1, 列表)



@acad.decorator_command
def llyx_rect_sidex_panjiayuan_for():
    zhu_ui_llyx_ness()
    objidlist = acad.SSGetIdList([[0, "LWPOLYLINE"]])
    with acad.transaction() as trans:
        acad.ChangeObjectIdLayer(objidlist, "图层1")
        result = acad.TransAutoFindRectPointList(objidlist)
        for pt1, pt2, dr1 in result:
            pt1, pt2, line3 = zhu_trans_llyx_att_w_line2(pt1, pt2, 3.5)
            pt1, pt2, line3 = zhu_trans_llyx_per_p_line3(pt1, pt2, dr1, 6)
            pt1, pt2, line3 = zhu_trans_llyx_per_p_line3(pt1, pt2, dr1, 8.5)
            line3.Layer = "0"



# length_set_list = [100, 35, 35, 40, 25, 25]
# @acad.decorator_command
# def llyx_rect_sidex_set_for():
#     objidlist = acad.SSGetIdList([[0, "LWPOLYLINE"]])
#     with acad.transaction() as trans:
#         acad.ChangeObjectIdLayer(objidlist, "图层1")
#         result = acad.TransAutoFindRectFencePointList(pb1, pb2, objidlist)
#         for pt1, pt2, dr1 in result:
#                 for m, length in enumerate(length_set_list):
#                     dr1 = acad.Vec3ResetLength(dr1, length)
#                     po1 = acad.Vec3Add(pt1, dr1)
#                     po2 = acad.Vec3Add(pt2, dr1)
#                     match m:
#                         case 1: 
#                             po1, po2 = acad.GetAttachWDirectPointList(po1, po2, length)
#                         case 3:
#                             po1 = acad.GetAttachNDirectPointPt1(po1, po2, length)
#                             po2 = acad.GetAttachWDirectPointPt2(po1, po2, length)
#                     line1 = acad.AddLine(pt1, po1)
#                     line2 = acad.AddLine(pt2, po2)
#                     line3 = acad.AddLine(po1, po2, "图层1")
#                     pt1, pt2 = po1, po2
#                 line3.Layer = "0"

# [acad.AddPoint(pt1) for pt1 in pline_point_list]
# pt1, pt2 = acad.GetBoundXY(ss1)
# acad.GetOSMODE()
# acad.AddRect(pt1, pt2)
# acad.SetOSMODE() 
# ss1 = acad.SSGet([[-4, "<OR"],[0, "LWPOLYLINE"],[0, "CIRCLE"], [-4, "OR>"]])

# TypedValue[] typeValue = {
#  new TypedValue((int) DxfCode.Operator, "<or"),
#  // 选择标准1
#  new TypedValue((int) DxfCode.Operator, "<and"),
#  new TypedValue((int) DxfCode.Start, "TEXT"),
#  new TypedValue((int) DxfCode.LayerName, "lay01"),
#  new TypedValue((int) DxfCode.Color, "1"),
#  new TypedValue((int) DxfCode.Operator, "and>"),
#  // 选择标准2
#  new TypedValue((int) DxfCode.Operator, "<and"),
#  new TypedValue((int) DxfCode.Start, "LWPOLYLINE"),
#  new TypedValue((int) DxfCode.Color, "5"),
#  new TypedValue((int) DxfCode.Operator, "<or"),
#  new TypedValue((int) DxfCode.LayerName, "lay02"),
#  new TypedValue((int) DxfCode.LayerName, "lay03"),
#  new TypedValue((int) DxfCode.Operator, "or>"),
#  new TypedValue((int) DxfCode.Operator, "and>"),
#  new TypedValue((int) DxfCode.Operator, "or>")
# };

# // 将筛选条件传到过滤器
# SelectionFilter selFiter = new SelectionFilter(typeValue);