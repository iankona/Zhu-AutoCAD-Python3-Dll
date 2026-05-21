import clr

import acad
import academit

import System

def 命令(): 
    academit.添加命令("lljf-line-sublength", lljf_line_sublength)
    academit.添加命令("lljf-line-subcount", lljf_line_subcount)
    academit.添加命令("lljf-line-gap-sublength", lljf_line_gap_sublength)
    academit.添加命令("lljf-line-gap-subcount", lljf_line_gap_subcount)
    academit.添加命令("lljf-line-clip-sublength",lljf_line_clip_sublength)
    academit.添加命令("lljf-line-clip-sublength-subcount",lljf_line_clip_sublength_subcount)
    academit.添加命令("lljf-curve-sublength", lljf_curve_sublength)
    academit.添加命令("lljf-curve-subcount", lljf_curve_subcount)
    academit.添加命令("lljf-curve-gap-sublength", lljf_curve_gap_sublength)
    academit.添加命令("lljf-curve-gap-subcount", lljf_curve_gap_subcount)
    academit.添加命令("lljf-curve-clip-sublength", lljf_curve_clip_sublength)
    academit.添加命令("lljf-curve-clip-sublength-subcount", lljf_curve_clip_sublength_subcount)
    academit.添加命令("lljf-line-sublength-set1", lljf_line_sublength_set1)



llzhu_jf_gapa = 0
llzhu_jf_gapb = 0
llzhu_jf_length_pipe = 30
llzhu_jf_length_sube = 300
llzhu_jf_length_numb = 5
llzhu_jf_length_pere = 300

def llzhu_ui_jf_input_clip_sublength():
    global llzhu_jf_gapa, llzhu_jf_gapb, llzhu_jf_length_pipe, llzhu_jf_length_sube, llzhu_jf_length_numb, llzhu_jf_length_pere
    gapa = acad.GetDouble(llzhu_jf_gapa, "请输入GapA:")
    pipe = acad.GetDouble(llzhu_jf_length_pipe, "请输入方管宽度:")
    sube = acad.GetDouble(llzhu_jf_length_sube, "请输入分段间隔:")
    pere = acad.GetDistance(llzhu_jf_length_pere, "请输入立杆长度:")
    if pipe != None: llzhu_jf_length_pipe = pipe
    if sube != None: llzhu_jf_length_sube = sube
    if pere != None: llzhu_jf_length_pere = pere
    llzhu_jf_gapa = gapa

def llzhu_ui_jf_input_clip_sublength_distance(distance):
    global llzhu_jf_gapa, llzhu_jf_gapb, llzhu_jf_length_pipe, llzhu_jf_length_sube, llzhu_jf_length_numb, llzhu_jf_length_pere
    gapa = llzhu_jf_gapa
    distance = distance - llzhu_jf_length_pipe
    if gapa <= 0:
        subcount = distance / (llzhu_jf_length_sube + llzhu_jf_length_pipe)
        count = int(subcount) # 取整
        llzhu_jf_length_numb = count
        gaplength = distance - (llzhu_jf_length_sube + llzhu_jf_length_pipe)*count # + llzhu_jf_length_pipe - llzhu_jf_length_pipe
        llzhu_jf_gapa, llzhu_jf_gapb = gaplength/2, gaplength/2
    else:
        distance = distance - llzhu_jf_gapa
        subcount = distance / (llzhu_jf_length_sube + llzhu_jf_length_pipe)
        count = int(subcount) # 四舍五入
        llzhu_jf_length_numb = count
        gaplength = distance - (llzhu_jf_length_sube + llzhu_jf_length_pipe)*count
        llzhu_jf_gapa, llzhu_jf_gapb = gapa, gaplength


def llzhu_ui_jf_input_clip_sublength_subcount():
    global llzhu_jf_gapa, llzhu_jf_gapb, llzhu_jf_length_pipe, llzhu_jf_length_sube, llzhu_jf_length_numb, llzhu_jf_length_pere
    gapa = acad.GetDouble(llzhu_jf_gapa, "请输入GapA:")
    pipe = acad.GetDouble(llzhu_jf_length_pipe, "请输入方管宽度:")
    sube = acad.GetDouble(llzhu_jf_length_sube, "请输入分段间隔:")
    numb = acad.GetInt(llzhu_jf_length_numb, "请输入间隔数量:")
    pere = acad.GetDistance(llzhu_jf_length_pere, "请输入立杆长度:")
    if pipe != None: llzhu_jf_length_pipe = pipe
    if sube != None: llzhu_jf_length_sube = sube
    if numb != None: llzhu_jf_length_numb = numb
    if pere != None: llzhu_jf_length_pere = pere
    llzhu_jf_gapa = gapa


def llzhu_ui_jf_input_clip_sublength_subcount_distance(distance):
    global llzhu_jf_gapa, llzhu_jf_gapb, llzhu_jf_length_pipe, llzhu_jf_length_sube, llzhu_jf_length_numb, llzhu_jf_length_pere
    gapa = llzhu_jf_gapa
    distance = distance - llzhu_jf_length_pipe
    gaplength = distance - (llzhu_jf_length_sube + llzhu_jf_length_pipe)*llzhu_jf_length_numb # + llzhu_jf_length_pipe - llzhu_jf_length_pipe
    if gapa <= 0 or gapa >= gaplength:
        llzhu_jf_gapa, llzhu_jf_gapb = gaplength/2, gaplength/2
    else:
        llzhu_jf_gapa, llzhu_jf_gapb = gapa, gaplength-gapa




def llzhu_ui_jf_input_gap_sublength():
    global llzhu_jf_gapa, llzhu_jf_gapb, llzhu_jf_length_pipe, llzhu_jf_length_sube, llzhu_jf_length_numb, llzhu_jf_length_pere
    gapa = acad.GetDouble(llzhu_jf_gapa, "请输入GapA:")
    gapb = acad.GetDouble(llzhu_jf_gapb, "请输入GapB:")
    pipe = acad.GetDouble(llzhu_jf_length_pipe, "请输入方管宽度:")
    sube = acad.GetDouble(llzhu_jf_length_sube, "请输入分段间隔:")
    pere = acad.GetDistance(llzhu_jf_length_pere, "请输入立杆长度:")
    if gapa != None: llzhu_jf_gapa = gapa
    if gapb != None: llzhu_jf_gapb = gapb
    if pipe != None: llzhu_jf_length_pipe = pipe
    if sube != None: llzhu_jf_length_sube = sube
    if pere != None: llzhu_jf_length_pere = pere

def llzhu_ui_jf_input_gap_sublength_distance(distance):
    global llzhu_jf_gapa, llzhu_jf_gapb, llzhu_jf_length_pipe, llzhu_jf_length_sube, llzhu_jf_length_numb, llzhu_jf_length_pere
    distance = distance - llzhu_jf_gapa - llzhu_jf_gapb
    distance = distance - llzhu_jf_length_pipe
    subcount = distance / (llzhu_jf_length_sube + llzhu_jf_length_pipe)
    count = round(subcount) # 四舍五入
    llzhu_jf_length_sube = distance/count - llzhu_jf_length_pipe
    llzhu_jf_length_numb = count


def llzhu_ui_jf_input_gap_subcount():
    global llzhu_jf_gapa, llzhu_jf_gapb, llzhu_jf_length_pipe, llzhu_jf_length_sube, llzhu_jf_length_numb, llzhu_jf_length_pere
    gapa = acad.GetDouble(llzhu_jf_gapa, "请输入GapA:")
    gapb = acad.GetDouble(llzhu_jf_gapb, "请输入GapB:")
    pipe = acad.GetDouble(llzhu_jf_length_pipe, "请输入方管宽度:")
    numb = acad.GetInt(llzhu_jf_length_numb, "请输入间隔数量:")
    pere = acad.GetDistance(llzhu_jf_length_pere, "请输入立杆长度:")
    if gapa != None: llzhu_jf_gapa = gapa
    if gapb != None: llzhu_jf_gapb = gapb
    if pipe != None: llzhu_jf_length_pipe = pipe
    if numb != None: llzhu_jf_length_numb = numb
    if pere != None: llzhu_jf_length_pere = pere

def llzhu_ui_jf_input_gap_subcount_distance(distance):
    global llzhu_jf_gapa, llzhu_jf_gapb, llzhu_jf_length_pipe, llzhu_jf_length_sube, llzhu_jf_length_numb, llzhu_jf_length_pere
    distance = distance - llzhu_jf_gapa - llzhu_jf_gapb
    distance = distance - llzhu_jf_length_pipe
    llzhu_jf_length_sube = distance/llzhu_jf_length_numb - llzhu_jf_length_pipe




def llzhu_ui_jf_input_sublength():
    global llzhu_jf_length_pipe, llzhu_jf_length_sube, llzhu_jf_length_numb, llzhu_jf_length_pere
    pipe = acad.GetDouble(llzhu_jf_length_pipe, "请输入方管宽度:")
    sube = acad.GetDouble(llzhu_jf_length_sube, "请输入分段间隔:")
    pere = acad.GetDistance(llzhu_jf_length_pere, "请输入立杆长度:")
    if pipe != None: llzhu_jf_length_pipe = pipe
    if sube != None: llzhu_jf_length_sube = sube
    if pere != None: llzhu_jf_length_pere = pere

def llzhu_ui_jf_input_sublength_distance(distance):
    global llzhu_jf_length_pipe, llzhu_jf_length_sube, llzhu_jf_length_numb, llzhu_jf_length_pere
    distance = distance + llzhu_jf_length_pipe
    subcount = distance / (llzhu_jf_length_sube + llzhu_jf_length_pipe)
    count = round(subcount) # 四舍五入
    llzhu_jf_length_sube = distance/count - llzhu_jf_length_pipe
    llzhu_jf_length_numb = count


def llzhu_ui_jf_input_subcount():
    global llzhu_jf_length_pipe, llzhu_jf_length_sube, llzhu_jf_length_numb, llzhu_jf_length_pere
    pipe = acad.GetDouble(llzhu_jf_length_pipe, "请输入方管宽度:")
    numb = acad.GetInt(llzhu_jf_length_numb, "请输入间隔数量:")
    pere = acad.GetDistance(llzhu_jf_length_pere, "请输入立杆长度:")
    if pipe != None: llzhu_jf_length_pipe = pipe
    if numb != None: llzhu_jf_length_numb = numb
    if pere != None: llzhu_jf_length_pere = pere

def llzhu_ui_jf_input_subcount_distance(distance):
    global llzhu_jf_length_pipe, llzhu_jf_length_sube, llzhu_jf_length_numb, llzhu_jf_length_pere
    distance = distance + llzhu_jf_length_pipe
    llzhu_jf_length_sube = distance/llzhu_jf_length_numb - llzhu_jf_length_pipe


def llzhu_calc_line_distance_pointlist(pt1, pt2, pt3):
    dr0 = acad.Direct(pt1, pt2)
    dr1 = acad.Vec3ResetLength(dr0, llzhu_jf_length_sube)
    if llzhu_jf_length_pipe > 0:
        dr2 = acad.Vec3ResetLength(dr0, llzhu_jf_length_pipe)
    per = acad.GetPerDirectResetLengthXY(pt1, pt2, pt3, llzhu_jf_length_pere)
    buflist = []
    if llzhu_jf_length_numb < 2: return []
    for i in range(llzhu_jf_length_numb-1):
        pt1 = acad.Vec3Add(pt1, dr1)
        po1, po2 = pt1, acad.Vec3Add(pt1, per)
        buflist.append([po1, po2])
        if llzhu_jf_length_pipe > 0:
            pt1 = acad.Vec3Add(pt1, dr2)
            po3, po4 = pt1, acad.Vec3Add(pt1, per)
            buflist.append([po3, po4])
    return buflist



def llzhu_calc_line_gap_distance_pointlist(pt1, pt2, pt3):
    dr0 = acad.Direct(pt1, pt2)
    dr1 = acad.Vec3ResetLength(dr0, llzhu_jf_length_sube)
    dr2 = acad.Vec3ResetLength(dr0, llzhu_jf_length_pipe)
    per = acad.GetPerDirectResetLengthXY(pt1, pt2, pt3, llzhu_jf_length_pere)
    if llzhu_jf_length_numb < 2: return []
    buflist = []
    dr3 = acad.Vec3ResetLength(dr0, llzhu_jf_gapa)
    pt1 = acad.Vec3Add(pt1, dr3)
    po1, po2 = pt1, acad.Vec3Add(pt1, per)    
    pt1 = acad.Vec3Add(pt1, dr2)
    po3, po4 = pt1, acad.Vec3Add(pt1, per)
    buflist.append([po1, po2])
    buflist.append([po3, po4])
    for i in range(llzhu_jf_length_numb):
        pt1 = acad.Vec3Add(pt1, dr1)
        po1, po2 = pt1, acad.Vec3Add(pt1, per)    
        pt1 = acad.Vec3Add(pt1, dr2)
        po3, po4 = pt1, acad.Vec3Add(pt1, per)
        buflist.append([po1, po2])
        buflist.append([po3, po4])
    return buflist


# match perflag:
#     case  1: dr1 =  [-y,  x,  z]
#     case -1: dr1 =  [ y, -x,  z]

def llzhu_trans_curve_distance_pointlist(objid, pt1):
    objref = acad.TransObjectForRead(objid)
    start, final = [objref.StartPoint.X, objref.StartPoint.Y, objref.StartPoint.Z], [objref.EndPoint.X, objref.EndPoint.Y, objref.EndPoint.Z]
    if pt1 == None:
        buflist = []
        if llzhu_jf_length_numb < 2: return []
        sumlength, sublength, pipehalf = 0, llzhu_jf_length_sube + llzhu_jf_length_pipe/2, llzhu_jf_length_pipe/2
        for i in range(llzhu_jf_length_numb-1):
            sumlength += sublength
            point = objref.GetPointAtDist(sumlength)
            direct = objref.GetFirstDerivative(point)
            po1 = [point.X, point.Y, point.Z]
            dr1 = [ direct.Y, -direct.X, direct.Z]
            dr2 = [-direct.Y,  direct.X, direct.Z]
            dr1 = acad.Vec3ResetLength(dr1, llzhu_jf_length_pere)
            dr2 = acad.Vec3ResetLength(dr2, llzhu_jf_length_pere)
            po2 = acad.Vec3Add(po1, dr1)
            po3 = acad.Vec3Add(po1, dr2)
            buflist.append([po1, po2])
            buflist.append([po1, po3])
            sumlength += pipehalf
        return buflist
    perdr1 = acad.GetPerDirectXY(start, final, pt1)
    buflist = []
    if llzhu_jf_length_numb < 2: return []
    sumlength, sublength, pipehalf = 0, llzhu_jf_length_sube + llzhu_jf_length_pipe/2, llzhu_jf_length_pipe/2
    for i in range(llzhu_jf_length_numb-1):
        sumlength += sublength
        point = objref.GetPointAtDist(sumlength)
        direct = objref.GetFirstDerivative(point)
        po1 = [point.X, point.Y, point.Z]
        dr1 = [ direct.Y, -direct.X, direct.Z]
        dr2 = [-direct.Y,  direct.X, direct.Z]
        if acad.AngleFromDotDr1Dr2(perdr1, dr1) < 90: pass
        if acad.AngleFromDotDr1Dr2(perdr1, dr2) < 90: dr1 = dr2 
        dr1 = acad.Vec3ResetLength(dr1, llzhu_jf_length_pere)
        po2 = acad.Vec3Add(po1, dr1)
        buflist.append([po1, po2])
        sumlength += pipehalf
    return buflist    




def llzhu_trans_curve_gap_distance_pointlist(objid, pt1):
    objref = acad.TransObjectForRead(objid)
    start, final = [objref.StartPoint.X, objref.StartPoint.Y, objref.StartPoint.Z], [objref.EndPoint.X, objref.EndPoint.Y, objref.EndPoint.Z]
    if pt1 == None:
        buflist = []
        if llzhu_jf_length_numb < 2: return []
        sumlength, sublength, pipehalf = 0, llzhu_jf_length_sube + llzhu_jf_length_pipe/2, llzhu_jf_length_pipe/2
        sumlength += (llzhu_jf_gapa+pipehalf)
        point = objref.GetPointAtDist(sumlength)
        direct = objref.GetFirstDerivative(point)
        po1 = [point.X, point.Y, point.Z]
        dr1 = [ direct.Y, -direct.X, direct.Z]
        dr2 = [-direct.Y,  direct.X, direct.Z]
        dr1 = acad.Vec3ResetLength(dr1, llzhu_jf_length_pere)
        dr2 = acad.Vec3ResetLength(dr2, llzhu_jf_length_pere)
        po2 = acad.Vec3Add(po1, dr1)
        po3 = acad.Vec3Add(po1, dr2)
        buflist.append([po1, po2])
        buflist.append([po1, po3])
        sumlength += pipehalf
        for i in range(llzhu_jf_length_numb):
            sumlength += sublength
            point = objref.GetPointAtDist(sumlength)
            direct = objref.GetFirstDerivative(point)
            po1 = [point.X, point.Y, point.Z]
            dr1 = [ direct.Y, -direct.X, direct.Z]
            dr2 = [-direct.Y,  direct.X, direct.Z]
            dr1 = acad.Vec3ResetLength(dr1, llzhu_jf_length_pere)
            dr2 = acad.Vec3ResetLength(dr2, llzhu_jf_length_pere)
            po2 = acad.Vec3Add(po1, dr1)
            po3 = acad.Vec3Add(po1, dr2)
            buflist.append([po1, po2])
            buflist.append([po1, po3])
            sumlength += pipehalf
        return buflist
    perdr1 = acad.GetPerDirectXY(start, final, pt1)
    buflist = []
    if llzhu_jf_length_numb < 2: return []
    sumlength, sublength, pipehalf = 0, llzhu_jf_length_sube + llzhu_jf_length_pipe/2, llzhu_jf_length_pipe/2
    sumlength += (llzhu_jf_gapa+pipehalf)
    point = objref.GetPointAtDist(sumlength)
    direct = objref.GetFirstDerivative(point)
    po1 = [point.X, point.Y, point.Z]
    dr1 = [ direct.Y, -direct.X, direct.Z]
    dr2 = [-direct.Y,  direct.X, direct.Z]
    if acad.AngleFromDotDr1Dr2(perdr1, dr1) < 90: pass
    if acad.AngleFromDotDr1Dr2(perdr1, dr2) < 90: dr1 = dr2 
    dr1 = acad.Vec3ResetLength(dr1, llzhu_jf_length_pere)
    po2 = acad.Vec3Add(po1, dr1)
    buflist.append([po1, po2])
    sumlength += pipehalf
    for i in range(llzhu_jf_length_numb):
        sumlength += sublength
        point = objref.GetPointAtDist(sumlength)
        direct = objref.GetFirstDerivative(point)
        po1 = [point.X, point.Y, point.Z]
        dr1 = [ direct.Y, -direct.X, direct.Z]
        dr2 = [-direct.Y,  direct.X, direct.Z]
        if acad.AngleFromDotDr1Dr2(perdr1, dr1) < 90: pass
        if acad.AngleFromDotDr1Dr2(perdr1, dr2) < 90: dr1 = dr2 
        dr1 = acad.Vec3ResetLength(dr1, llzhu_jf_length_pere)
        po2 = acad.Vec3Add(po1, dr1)
        buflist.append([po1, po2])
        sumlength += pipehalf
    return buflist    









# StartParam
# EndParam
# StartPoint
# EndPoint
# GetPointAtParameter(double value)
# GetParameterAtPoint(Point3d point)
# GetDistanceAtParameter(double value)
# GetParameterAtDistance(double dist)
# GetDistAtPoint(Point3d point)
# GetPointAtDist(double value)
# GetFirstDerivative(Point3d point)
# GetFirstDerivative(double value)
# GetSecondDerivative(Point3d point)
# GetSecondDerivative(double value)


@acad.decorator_command
def lljf_line_sublength():
    llzhu_ui_jf_input_sublength()
    while True:
        pt1, pt2, pt3 = acad.GetPoint3() 
        if pt1 == None: return
        llzhu_ui_jf_input_sublength_distance(acad.Distance(pt1, pt2))
        result = llzhu_calc_line_distance_pointlist(pt1, pt2, pt3)
        with acad.transaction() as trans:
            for po1, po2 in result: acad.AddLine(po1, po2)

@acad.decorator_command
def lljf_line_subcount():
    llzhu_ui_jf_input_subcount()
    while True:
        pt1, pt2, pt3 = acad.GetPoint3() 
        if pt1 == None: return
        llzhu_ui_jf_input_subcount_distance(acad.Distance(pt1, pt2))
        result = llzhu_calc_line_distance_pointlist(pt1, pt2, pt3)
        with acad.transaction() as trans:
            for po1, po2 in result: acad.AddLine(po1, po2)


@acad.decorator_command
def lljf_line_gap_sublength():
    llzhu_ui_jf_input_gap_sublength()
    while True:
        pt1, pt2, pt3 = acad.GetPoint3() 
        if pt1 == None: return
        llzhu_ui_jf_input_gap_sublength_distance(acad.Distance(pt1, pt2))
        result = llzhu_calc_line_gap_distance_pointlist(pt1, pt2, pt3)
        with acad.transaction() as trans:
            for po1, po2 in result: acad.AddLine(po1, po2)

@acad.decorator_command
def lljf_line_gap_subcount():
    llzhu_ui_jf_input_gap_subcount()
    while True:
        pt1, pt2, pt3 = acad.GetPoint3() 
        if pt1 == None: return
        llzhu_ui_jf_input_gap_subcount_distance(acad.Distance(pt1, pt2))
        result = llzhu_calc_line_gap_distance_pointlist(pt1, pt2, pt3)
        with acad.transaction() as trans:
            for po1, po2 in result: acad.AddLine(po1, po2)


@acad.decorator_command
def lljf_curve_sublength():
    llzhu_ui_jf_input_sublength()
    while True:
        objid = acad.EntSel()
        if acad.IsNoneObjectId(objid): return 
        length = acad.GetEntityLength(objid)
        llzhu_ui_jf_input_sublength_distance(length)
        pt1 = acad.GetPoint("请点击方向点: ")
        pipehalf = llzhu_jf_length_pipe/2
        with acad.transaction() as trans:
            result = llzhu_trans_curve_distance_pointlist(objid, pt1)
            for po1, po2 in result: 
                line = acad.DBObjectLine(po1, po2)
                collect = line.GetOffsetCurves(pipehalf)
                for objref in collect: acad.AddDBObject(objref)
                collect = line.GetOffsetCurves(-pipehalf)
                for objref in collect: acad.AddDBObject(objref)


@acad.decorator_command
def lljf_curve_subcount():
    llzhu_ui_jf_input_subcount()
    while True:
        objid = acad.EntSel()
        if acad.IsNoneObjectId(objid): return 
        length = acad.GetEntityLength(objid)
        llzhu_ui_jf_input_subcount_distance(length)
        pt1 = acad.GetPoint("请点击方向点: ")
        pipehalf = llzhu_jf_length_pipe/2
        with acad.transaction() as trans:
            result = llzhu_trans_curve_distance_pointlist(objid, pt1)
            for po1, po2 in result: 
                # acad.AddLine(po1, po2)
                line = acad.DBObjectLine(po1, po2)
                collect = line.GetOffsetCurves(pipehalf)
                for objref in collect: acad.AddDBObject(objref)
                collect = line.GetOffsetCurves(-pipehalf)
                for objref in collect: acad.AddDBObject(objref)




@acad.decorator_command
def lljf_curve_gap_sublength():
    llzhu_ui_jf_input_gap_sublength()
    while True:
        objid = acad.EntSel()
        if acad.IsNoneObjectId(objid): return 
        length = acad.GetEntityLength(objid)
        llzhu_ui_jf_input_gap_sublength_distance(length)
        pt1 = acad.GetPoint("请点击方向点: ")
        pipehalf = llzhu_jf_length_pipe/2
        with acad.transaction() as trans:
            result = llzhu_trans_curve_gap_distance_pointlist(objid, pt1)
            for po1, po2 in result: 
                line = acad.DBObjectLine(po1, po2)
                collect = line.GetOffsetCurves(pipehalf)
                for objref in collect: acad.AddDBObject(objref)
                collect = line.GetOffsetCurves(-pipehalf)
                for objref in collect: acad.AddDBObject(objref)


@acad.decorator_command
def lljf_curve_gap_subcount():
    llzhu_ui_jf_input_gap_subcount()
    while True:
        objid = acad.EntSel()
        if acad.IsNoneObjectId(objid): return 
        length = acad.GetEntityLength(objid)
        llzhu_ui_jf_input_gap_subcount_distance(length)
        pt1 = acad.GetPoint("请点击方向点: ")
        pipehalf = llzhu_jf_length_pipe/2
        with acad.transaction() as trans:
            result = llzhu_trans_curve_gap_distance_pointlist(objid, pt1)
            for po1, po2 in result: 
                # acad.AddLine(po1, po2)
                line = acad.DBObjectLine(po1, po2)
                collect = line.GetOffsetCurves(pipehalf)
                for objref in collect: acad.AddDBObject(objref)
                collect = line.GetOffsetCurves(-pipehalf)
                for objref in collect: acad.AddDBObject(objref)



@acad.decorator_command
def lljf_line_clip_sublength():
    llzhu_ui_jf_input_clip_sublength()
    while True:
        pt1, pt2, pt3 = acad.GetPoint3() 
        if pt1 == None: return
        llzhu_ui_jf_input_clip_sublength_distance(acad.Distance(pt1, pt2))
        result = llzhu_calc_line_gap_distance_pointlist(pt1, pt2, pt3)
        with acad.transaction() as trans:
            for po1, po2 in result: acad.AddLine(po1, po2)

@acad.decorator_command
def lljf_line_clip_sublength_subcount():
    llzhu_ui_jf_input_clip_sublength_subcount()
    while True:
        pt1, pt2, pt3 = acad.GetPoint3() 
        if pt1 == None: return
        llzhu_ui_jf_input_clip_sublength_subcount_distance(acad.Distance(pt1, pt2))
        result = llzhu_calc_line_gap_distance_pointlist(pt1, pt2, pt3)
        with acad.transaction() as trans:
            for po1, po2 in result: acad.AddLine(po1, po2)


@acad.decorator_command
def lljf_curve_clip_sublength():
    llzhu_ui_jf_input_clip_sublength()
    while True:
        objid = acad.EntSel()
        if acad.IsNoneObjectId(objid): return 
        length = acad.GetEntityLength(objid)
        llzhu_ui_jf_input_clip_sublength_distance(length)
        pt1 = acad.GetPoint("请点击方向点: ")
        pipehalf = llzhu_jf_length_pipe/2
        with acad.transaction() as trans:
            result = llzhu_trans_curve_gap_distance_pointlist(objid, pt1)
            for po1, po2 in result: 
                line = acad.DBObjectLine(po1, po2)
                collect = line.GetOffsetCurves(pipehalf)
                for objref in collect: acad.AddDBObject(objref)
                collect = line.GetOffsetCurves(-pipehalf)
                for objref in collect: acad.AddDBObject(objref)


@acad.decorator_command
def lljf_curve_clip_sublength_subcount():
    llzhu_ui_jf_input_clip_sublength_subcount()
    while True:
        objid = acad.EntSel()
        if acad.IsNoneObjectId(objid): return 
        length = acad.GetEntityLength(objid)
        llzhu_ui_jf_input_clip_sublength_subcount_distance(length)
        pt1 = acad.GetPoint("请点击方向点: ")
        pipehalf = llzhu_jf_length_pipe/2
        with acad.transaction() as trans:
            result = llzhu_trans_curve_gap_distance_pointlist(objid, pt1)
            for po1, po2 in result: 
                line = acad.DBObjectLine(po1, po2)
                collect = line.GetOffsetCurves(pipehalf)
                for objref in collect: acad.AddDBObject(objref)
                collect = line.GetOffsetCurves(-pipehalf)
                for objref in collect: acad.AddDBObject(objref)


@acad.decorator_command
def lljf_line_sublength_set1():
    pipe1 = acad.GetDouble(40, "请输入方管1宽度:")
    pipe2 = acad.GetDouble(20, "请输入方管2宽度:")
    sublength = acad.GetDouble(70, "请输入分段间隔:")
    perlength = acad.GetDistance(300, "请输入立杆长度:")
    while True:
        pt1, pt2, pt3 = acad.GetPoint3() 
        if pt1 == None: return
        distance = acad.Distance(pt1, pt2)
        count = 1
        lastbias = distance
        buflist = []
        while True:
            count += 1
            acount = count
            bcount = count - 1
            ccount = acount + bcount + 1
            remain = distance - acount*pipe1 - bcount*pipe2
            length = remain / ccount
            bias = abs(length-sublength)
            buflist.append([bias, acount, bcount, ccount, length])
            # print([bias, acount, bcount, ccount, length])
            if bias < lastbias: lastbias = bias
            if bias > lastbias: break
        bias, acount, bcount, ccount, sublength = buflist[-2]

        per = acad.GetPerDirectResetLengthXY(pt1, pt2, pt3, perlength)
        dr0 = acad.Direct(pt1, pt2)
        dr1 = acad.Vec3ResetLength(dr0, sublength)
        dr2 = acad.Vec3ResetLength(dr0, pipe1)
        dr3 = acad.Vec3ResetLength(dr0, sublength)
        dr4 = acad.Vec3ResetLength(dr0, pipe2)
        ptlist1 = []
        for i in range(bcount):
            pt1 = acad.Vec3Add(pt1, dr1)
            ptlist1.append(pt1)
            pt1 = acad.Vec3Add(pt1, dr2)
            ptlist1.append(pt1)
            pt1 = acad.Vec3Add(pt1, dr3)
            ptlist1.append(pt1)
            pt1 = acad.Vec3Add(pt1, dr4)
            ptlist1.append(pt1)
        pt1 = acad.Vec3Add(pt1, dr1)
        ptlist1.append(pt1)
        pt1 = acad.Vec3Add(pt1, dr2)
        ptlist1.append(pt1)

        ptlist2 = [acad.Vec3Add(po1, per) for po1 in ptlist1]

        with acad.transaction() as trans:
            for po1, po2 in zip(ptlist1, ptlist2): acad.AddLine(po1, po2)


# pipe1 = acad.GetDouble(llzhu_jf_length_pipe, "请输入方管1宽度:")
# pipe2 = acad.GetDouble(llzhu_jf_length_pipe, "请输入方管2宽度:")
# sube = acad.GetDouble(llzhu_jf_length_sube, "请输入分段间隔:")
# numb = acad.GetInt(llzhu_jf_length_numb, "请输入间隔数量:")
# pere = acad.GetDistance(llzhu_jf_length_pere, "请输入立杆长度:")



