import clr

import acad
import academit

import System

def 命令(): 
    # academit.添加命令("llyx", llyx)
    academit.添加命令("llyx-side3", llyx_side3)
    academit.添加命令("llyx-sidexfor", llyx_sidexfor)
    academit.添加命令("llyx-sidex-nfor", llyx_sidex_nfor)
    academit.添加命令("llyx-sidex-wfor", llyx_sidex_wfor)

    academit.添加命令("llyx-rec-side3for", llyx_rec_side3for)
    academit.添加命令("llyx-rec-sidexfor", llyx_rec_sidexfor)
    academit.添加命令("llyx-rec-sidex-nfor", llyx_rec_sidex_nfor)
    academit.添加命令("llyx-rec-sidex-wfor", llyx_rec_sidex_wfor)
    academit.添加命令("llyx-rec-sidex-setfor", llyx_rec_sidex_setfor)
    pass

offset_length = 50

@acad.decorator_command
def llyx():
    global offset_length
    length = acad.GetDouble(offset_length, "请输入偏移大小: ")
    if length != None: offset_length = length
    pt1, pt2, pt3 = acad.GetPoint3("请选择第1个顶点: ", "请选择第2个顶点: ", "请点击方向顶点: ")
    if pt1 == None: return
    dr1 = acad.GetPerDirectResetLengthXY(pt1, pt2, pt3, offset_length) # 点在线上，WhichSideOfLineXY 会出现cannot access local variable 'result' where it is not associated with a value
    po1 = acad.Vec3Add(pt1, dr1)
    po2 = acad.Vec3Add(pt2, dr1)
    with acad.transaction() as trans:
        acad.AddLine(po1, po2)


@acad.decorator_command_undo
def llyx_side3():
    global offset_length
    length = acad.GetDouble(offset_length, "请输入偏移大小: ")
    if length != None: offset_length = length
    pt1, pt2, pt3 = acad.GetPoint3("请选择第1个顶点: ", "请选择第2个顶点: ", "请点击方向顶点: ")
    if pt1 == None: return
    dr1 = acad.GetPerDirectResetLengthXY(pt1, pt2, pt3, offset_length) # 点在线上，WhichSideOfLineXY 会出现cannot access local variable 'result' where it is not associated with a value
    po1 = acad.Vec3Add(pt1, dr1)
    po2 = acad.Vec3Add(pt2, dr1)
    with acad.transaction() as trans:
        acad.AddLine(pt1, po1)
        acad.AddLine(pt2, po2)
        acad.AddLine(po1, po2)



@acad.decorator_command
def llyx_sidexfor():
    列表 = acad.GetDoubleListLimitCount()
    if 列表 == None: return
    while True:
        pt1, pt2, pt3 = acad.GetPoint3("请选择第1个顶点: ", "请选择第2个顶点: ", "请点击方向顶点: ")
        if pt1 == None: break
        with acad.transaction() as trans:
            dr1 = acad.GetPerDirectXY(pt1, pt2, pt3)
            for length in 列表:
                dr1 = acad.Vec3ResetLength(dr1, length)
                po1 = acad.Vec3Add(pt1, dr1)
                po2 = acad.Vec3Add(pt2, dr1)
                line1 = acad.AddLine(pt1, po1)
                line2 = acad.AddLine(pt2, po2)
                line3 = acad.AddLine(po1, po2, "图层1")
                pt1, pt2 = po1, po2
            line3.Layer = "0"

@acad.decorator_command
def llyx_sidex_nfor():
    列表 = acad.GetDoubleListLimitCount()
    if 列表 == None: return
    while True:
        pt1, pt2, pt3 = acad.GetPoint3("请选择第1个顶点: ", "请选择第2个顶点: ", "请点击方向顶点: ")
        if pt1 == None: break
        with acad.transaction() as trans:
            dr1 = acad.GetPerDirectXY(pt1, pt2, pt3)
            for m, length in enumerate(列表):
                dr1 = acad.Vec3ResetLength(dr1, length)
                po1 = acad.Vec3Add(pt1, dr1)
                po2 = acad.Vec3Add(pt2, dr1)
                if m == 1: po1, po2 = acad.GetAttachNDirectPointList(po1, po2, length)
                line1 = acad.AddLine(pt1, po1)
                line2 = acad.AddLine(pt2, po2)
                line3 = acad.AddLine(po1, po2, "图层1")
                pt1, pt2 = po1, po2
            line3.Layer = "0"

@acad.decorator_command
def llyx_sidex_wfor():
    列表 = acad.GetDoubleListLimitCount()
    if 列表 == None: return
    while True:
        pt1, pt2, pt3 = acad.GetPoint3("请选择第1个顶点: ", "请选择第2个顶点: ", "请点击方向顶点: ")
        if pt1 == None: break
        with acad.transaction() as trans:
            dr1 = acad.GetPerDirectXY(pt1, pt2, pt3)
            for m, length in enumerate(列表):
                dr1 = acad.Vec3ResetLength(dr1, length)
                po1 = acad.Vec3Add(pt1, dr1)
                po2 = acad.Vec3Add(pt2, dr1)
                if m == 1: po1, po2 = acad.GetAttachWDirectPointList(po1, po2, length)
                line1 = acad.AddLine(pt1, po1)
                line2 = acad.AddLine(pt2, po2)
                line3 = acad.AddLine(po1, po2, "图层1")
                pt1, pt2 = po1, po2
            line3.Layer = "0"


@acad.decorator_command
def llyx_rec_side3for():
    global offset_length
    length = acad.GetDouble(offset_length, "请输入偏移大小: ")
    if length != None: offset_length = length
    objidlist = acad.SSGetIdList([[0, "LWPOLYLINE"]])
    with acad.transaction() as trans:
        acad.ChangeObjectIdLayer(objidlist, "图层1")
        for objid in objidlist:
            center = acad.GetEntityBoundCenterXY(objid)
            pline_point_list = acad.GetLWPolyLinePointList(objid)
            for i in range(len(pline_point_list)-1):
                pt1 = pline_point_list[i]
                pt2 = pline_point_list[i+1]
                pt3 = center
                perflag = acad.GetPerflagXY(pt1, pt2, pt3)
                perflag = -perflag
                dr1 = acad.GetPerDirectWithPerflagResetLengthXY(pt1, pt2, perflag, offset_length) 
                po1 = acad.Vec3Add(pt1, dr1)
                po2 = acad.Vec3Add(pt2, dr1)
                acad.AddLine(pt1, po1)
                acad.AddLine(pt2, po2)
                acad.AddLine(po1, po2)


@acad.decorator_command
def llyx_rec_sidexfor():
    列表 = acad.GetDoubleListLimitCount()
    if 列表 == None: return
    objidlist = acad.SSGetIdList([[0, "LWPOLYLINE"]])
    with acad.transaction() as trans:
        acad.ChangeObjectIdLayer(objidlist, "图层1")
        for objid in objidlist:
            center = acad.GetEntityBoundCenterXY(objid)
            pline_point_list = acad.GetLWPolyLinePointList(objid)
            for i in range(len(pline_point_list)-1):
                pt1 = pline_point_list[i]
                pt2 = pline_point_list[i+1]
                pt3 = center
                perflag = acad.GetPerflagXY(pt1, pt2, pt3)
                perflag = -perflag
                dr1 = acad.GetPerDirectWithPerflagXY(pt1, pt2, perflag) 
                for length in 列表:
                    dr1 = acad.Vec3ResetLength(dr1, length)
                    po1 = acad.Vec3Add(pt1, dr1)
                    po2 = acad.Vec3Add(pt2, dr1)                  
                    line1 = acad.AddLine(pt1, po1)
                    line2 = acad.AddLine(pt2, po2)
                    line3 = acad.AddLine(po1, po2, "图层1")
                    pt1, pt2 = po1, po2
                line3.Layer = "0"

@acad.decorator_command
def llyx_rec_sidex_nfor():
    列表 = acad.GetDoubleListLimitCount()
    if 列表 == []: return
    objidlist = acad.SSGetIdList([[0, "LWPOLYLINE"]])
    with acad.transaction() as trans:
        acad.ChangeObjectIdLayer(objidlist, "图层1")
        for objid in objidlist:
            center = acad.GetEntityBoundCenterXY(objid)
            pline_point_list = acad.GetLWPolyLinePointList(objid)
            for i in range(len(pline_point_list)-1):
                pt1 = pline_point_list[i]
                pt2 = pline_point_list[i+1]
                pt3 = center
                perflag = acad.GetPerflagXY(pt1, pt2, pt3)
                perflag = -perflag
                dr1 = acad.GetPerDirectWithPerflagXY(pt1, pt2, perflag) 
                for m, length in enumerate(列表):
                    dr1 = acad.Vec3ResetLength(dr1, length)
                    po1 = acad.Vec3Add(pt1, dr1)
                    po2 = acad.Vec3Add(pt2, dr1)
                    if m == 1: po1, po2 = acad.GetAttachNDirectPointList(po1, po2, length)
                    line1 = acad.AddLine(pt1, po1)
                    line2 = acad.AddLine(pt2, po2)
                    line3 = acad.AddLine(po1, po2, "图层1")
                    pt1, pt2 = po1, po2
                line3.Layer = "0"


@acad.decorator_command
def llyx_rec_sidex_wfor():
    列表 = acad.GetDoubleListLimitCount()
    if 列表 == None: return
    objidlist = acad.SSGetIdList([[0, "LWPOLYLINE"]])
    with acad.transaction() as trans:
        acad.ChangeObjectIdLayer(objidlist, "图层1")
        for objid in objidlist:
            center = acad.GetEntityBoundCenterXY(objid)
            pline_point_list = acad.GetLWPolyLinePointList(objid)
            for i in range(len(pline_point_list)-1):
                pt1 = pline_point_list[i]
                pt2 = pline_point_list[i+1]
                pt3 = center
                perflag = acad.GetPerflagXY(pt1, pt2, pt3)
                perflag = -perflag
                dr1 = acad.GetPerDirectWithPerflagXY(pt1, pt2, perflag) 
                for m, length in enumerate(列表):
                    dr1 = acad.Vec3ResetLength(dr1, length)
                    po1 = acad.Vec3Add(pt1, dr1)
                    po2 = acad.Vec3Add(pt2, dr1)
                    if m == 1: po1, po2 = acad.GetAttachWDirectPointList(po1, po2, length)
                    line1 = acad.AddLine(pt1, po1)
                    line2 = acad.AddLine(pt2, po2)
                    line3 = acad.AddLine(po1, po2, "图层1")
                    pt1, pt2 = po1, po2
                line3.Layer = "0"


length_set_list = [100, 35, 35, 40, 25, 25]
@acad.decorator_command
def llyx_rec_sidex_setfor():
    objidlist = acad.SSGetIdList([[0, "LWPOLYLINE"]])
    with acad.transaction() as trans:
        acad.ChangeObjectIdLayer(objidlist, "图层1")
        for objid in objidlist:
            center = acad.GetEntityBoundCenterXY(objid)
            pline_point_list = acad.GetLWPolyLinePointList(objid)
            for i in range(len(pline_point_list)-1):
                pt1 = pline_point_list[i]
                pt2 = pline_point_list[i+1]
                pt3 = center
                perflag = acad.GetPerflagXY(pt1, pt2, pt3)
                perflag = -perflag
                dr1 = acad.GetPerDirectWithPerflagXY(pt1, pt2, perflag) 
                for m, length in enumerate(length_set_list):
                    dr1 = acad.Vec3ResetLength(dr1, length)
                    po1 = acad.Vec3Add(pt1, dr1)
                    po2 = acad.Vec3Add(pt2, dr1)
                    match m:
                        case 1: 
                            po1, po2 = acad.GetAttachWDirectPointList(po1, po2, length)
                        case 3:
                            po1 = acad.GetAttachNDirectPointPt1(po1, po2, length)
                            po2 = acad.GetAttachWDirectPointPt2(po1, po2, length)
                    line1 = acad.AddLine(pt1, po1)
                    line2 = acad.AddLine(pt2, po2)
                    line3 = acad.AddLine(po1, po2, "图层1")
                    pt1, pt2 = po1, po2
                line3.Layer = "0"


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