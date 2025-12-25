import clr

import acad
import academit

import System

def 命令(): 
    academit.添加命令("llyx", llyx)
    academit.添加命令("llyxfor", llyxfor)
    academit.添加命令("llyx-side3", llyx_side3)
    academit.添加命令("llyx-sidexfor", llyx_sidexfor)
    academit.添加命令("llyx-sidex-nfor", llyx_sidex_nfor)
    academit.添加命令("llyx-sidex-wfor", llyx_sidex_wfor)

    academit.添加命令("llyx-rec-side3for", llyx_rec_side3for)
    academit.添加命令("llyx-rec-sidexfor", llyx_rec_sidexfor)
    pass

offset_length = 50

def llyx():
    global offset_length
    acad.GetActiveDocument()
    try:
        # acad.GetUndo()
        length = acad.GetDouble(offset_length, "请输入偏移大小: ")
        if length != None: offset_length = length
        pt1 = acad.GetPoint("请选择第1个顶点: ")
        if pt1 == None: return
        pt2 = acad.GetPoint("请选择第2个顶点: ")
        if pt2 == None: return
        pt3 = acad.GetPoint("请点击方向顶点: ")
        if pt3 == None: return
        dr1 = acad.GetPerDirectResetLengthXY(pt1, pt2, pt3, offset_length) # 点在线上，WhichSideOfLineXY 会出现cannot access local variable 'result' where it is not associated with a value
        po1 = acad.Vec3Add(pt1, dr1)
        po2 = acad.Vec3Add(pt2, dr1)
        acad.GetOSMODE()
        acad.AddLine(po1, po2)
        acad.SetOSMODE()
        # acad.SetUndo()
    except Exception as e:
        # acad.HappenErrorUndo()
        print(e)
        acad.Prompt("...函数出错...\n")

    
def llyxfor():
    global offset_length
    acad.GetActiveDocument()
    try:
        # acad.GetUndo()
        length = acad.GetDouble(offset_length, "请输入偏移大小: ")
        if length != None: offset_length = length
        while True:
            pt1 = acad.GetPoint("请选择第1个顶点: ")
            if pt1 == None: break
            pt2 = acad.GetPoint("请选择第2个顶点: ")
            if pt2 == None: break
            pt3 = acad.GetPoint("请点击方向顶点: ")
            if pt3 == None: break
            dr1 = acad.GetPerDirectResetLengthXY(pt1, pt2, pt3, offset_length) # 点在线上，WhichSideOfLineXY 会出现cannot access local variable 'result' where it is not associated with a value
            po1 = acad.Vec3Add(pt1, dr1)
            po2 = acad.Vec3Add(pt2, dr1)
            acad.GetOSMODE()
            acad.AddLine(po1, po2)
            acad.SetOSMODE()
        # acad.SetUndo()
    except Exception as e:
        # acad.HappenErrorUndo()
        print(e)
        acad.Prompt("...函数出错...\n")



def llyx_side3():
    global offset_length
    acad.GetActiveDocument()
    try:
        acad.GetUndo()
        length = acad.GetDouble(offset_length, "请输入偏移大小: ")
        if length != None: offset_length = length
        pt1 = acad.GetPoint("请选择第1个顶点: ")
        if pt1 == None: return
        pt2 = acad.GetPoint("请选择第2个顶点: ")
        if pt2 == None: return
        pt3 = acad.GetPoint("请点击方向顶点: ")
        if pt3 == None: return
        dr1 = acad.GetPerDirectResetLengthXY(pt1, pt2, pt3, offset_length) # 点在线上，WhichSideOfLineXY 会出现cannot access local variable 'result' where it is not associated with a value
        po1 = acad.Vec3Add(pt1, dr1)
        po2 = acad.Vec3Add(pt2, dr1)
        acad.GetOSMODE()
        acad.AddLine(pt1, po1)
        acad.AddLine(pt2, po2)
        acad.AddLine(po1, po2)
        acad.SetOSMODE()
        acad.SetUndo()
    except Exception as e:
        acad.HappenErrorUndo()
        print(e)
        acad.Prompt("...函数出错...\n")

def llyx_sidexfor():
    acad.GetActiveDocument()
    try:
        列表 = acad.GetDoubleListLimitCount()
        if 列表 == None: return
        while True:
            acad.GetUndo()
            pt1 = acad.GetPoint("请选择第1个顶点: ")
            if pt1 == None: break
            pt2 = acad.GetPoint("请选择第2个顶点: ")
            if pt2 == None: break
            pt3 = acad.GetPoint("请点击方向顶点: ")
            if pt3 == None: break
            acad.GetOSMODE()
            dr1 = acad.GetPerDirectXY(pt1, pt2, pt3)
            for length in 列表:
                dr1 = acad.Vec3ResetLength(dr1, length) # 点在线上，WhichSideOfLineXY 会出现cannot access local variable 'result' where it is not associated with a value
                po1 = acad.Vec3Add(pt1, dr1)
                po2 = acad.Vec3Add(pt2, dr1)
                acad.AddLine(pt1, po1)
                acad.AddLine(pt2, po2)
                acad.AddLine(po1, po2)
                pt1, pt2 = po1, po2
            acad.SetOSMODE()
            acad.SetUndo()
    except Exception as e:
        acad.HappenErrorUndo()
        print(e)
        acad.Prompt("...函数出错...\n")

def llyx_sidex_nfor():
    acad.GetActiveDocument()
    try:
        列表 = acad.GetDoubleListLimitCount()
        if 列表 == None: return
        while True:
            acad.GetUndo()
            pt1 = acad.GetPoint("请选择第1个顶点: ")
            if pt1 == None: break
            pt2 = acad.GetPoint("请选择第2个顶点: ")
            if pt2 == None: break
            pt3 = acad.GetPoint("请点击方向顶点: ")
            if pt3 == None: break
            acad.GetOSMODE()
            dr1 = acad.GetPerDirectXY(pt1, pt2, pt3)
            for i, length in enumerate(列表):
                dr1 = acad.Vec3ResetLength(dr1, length) # 点在线上，WhichSideOfLineXY 会出现cannot access local variable 'result' where it is not associated with a value
                po1 = acad.Vec3Add(pt1, dr1)
                po2 = acad.Vec3Add(pt2, dr1)
                if i == 1: po1, po2 = acad.GetAttachNDirectPoints(po1, po2, length)
                acad.AddLine(pt1, po1)
                acad.AddLine(pt2, po2)
                acad.AddLine(po1, po2)
                pt1, pt2 = po1, po2
            acad.SetOSMODE()
            acad.SetUndo()
    except Exception as e:
        acad.HappenErrorUndo()
        print(e)
        acad.Prompt("...函数出错...\n")


def llyx_sidex_wfor():
    acad.GetActiveDocument()
    try:
        列表 = acad.GetDoubleListLimitCount()
        if 列表 == None: return
        while True:
            acad.GetUndo()
            pt1 = acad.GetPoint("请选择第1个顶点: ")
            if pt1 == None: break
            pt2 = acad.GetPoint("请选择第2个顶点: ")
            if pt2 == None: break
            pt3 = acad.GetPoint("请点击方向顶点: ")
            if pt3 == None: break
            acad.GetOSMODE()
            dr1 = acad.GetPerDirectXY(pt1, pt2, pt3)
            for i, length in enumerate(列表):
                dr1 = acad.Vec3ResetLength(dr1, length) # 点在线上，WhichSideOfLineXY 会出现cannot access local variable 'result' where it is not associated with a value
                po1 = acad.Vec3Add(pt1, dr1)
                po2 = acad.Vec3Add(pt2, dr1)
                if i == 1: po1, po2 = acad.GetAttachWDirectPoints(po1, po2, length)
                acad.AddLine(pt1, po1)
                acad.AddLine(pt2, po2)
                acad.AddLine(po1, po2)
                pt1, pt2 = po1, po2
            acad.SetOSMODE()
            acad.SetUndo()
    except Exception as e:
        acad.HappenErrorUndo()
        print(e)
        acad.Prompt("...函数出错...\n")


def llyx_rec_side3for():
    global offset_length
    acad.GetActiveDocument()
    try:
        acad.GetUndo()
        length = acad.GetDouble(offset_length, "请输入偏移大小: ")
        if length != None: offset_length = length
        # ss1 = acad.SSGet([[-4, "<OR"],[0, "LWPOLYLINE"],[0, "CIRCLE"], [-4, "OR>"]])
        ss1 = acad.SSGet([[0, "LWPOLYLINE"]])
        # pt1, pt2 = acad.GetBoundXY(ss1)
        # acad.GetOSMODE()
        # acad.AddRect(pt1, pt2)
        # acad.SetOSMODE() 
        acad.GetOSMODE()   
        for objid in ss1.GetObjectIds():
            center = acad.GetEntityBoundCenterXY(objid)
            pline_point_list = acad.GetLWPolyLinePointList(objid)
            # [acad.AddPoint(pt1) for pt1 in pline_point_list]
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
        acad.SetOSMODE()
        acad.SetUndo()
    except Exception as e:
        acad.HappenErrorUndo()
        print(e)
        acad.Prompt("...函数出错...\n")


def llyx_rec_sidexfor():
    global offset_length
    acad.GetActiveDocument()
    try:
        acad.GetUndo()
        列表 = acad.GetDoubleListLimitCount()
        if 列表 == None: return
        ss1 = acad.SSGet([[0, "LWPOLYLINE"]])
        if ss1 == None: return
        acad.GetOSMODE()   
        for objid in ss1.GetObjectIds():
            center = acad.GetEntityBoundCenterXY(objid)
            pline_point_list = acad.GetLWPolyLinePointList(objid)
            # [acad.AddPoint(pt1) for pt1 in pline_point_list]
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
                    acad.AddLine(pt1, po1)
                    acad.AddLine(pt2, po2)
                    acad.AddLine(po1, po2)
                    pt1, pt2 = po1, po2
        acad.SetOSMODE()
        acad.SetUndo()
    except Exception as e:
        # acad.HappenErrorUndo()
        print(e)
        acad.Prompt("...函数出错...\n")

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