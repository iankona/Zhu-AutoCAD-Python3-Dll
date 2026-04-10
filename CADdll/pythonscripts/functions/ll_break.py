  
import acad
import academit
import System

def 命令(): 
    academit.添加命令("llccb", llccb)
    academit.添加命令("llbreak-all", llbreak_all)
    academit.添加命令("llbreak-point", llbreak_point)
    academit.添加命令("llbreak-subcount-for", llbreak_subcount_for)
    academit.添加命令("llbreak-sublength-for", llbreak_sublength_for)
    academit.添加命令("llbreak-line-use-select", llbreak_line_use_select)
    academit.添加命令("llbreak-select-use-line", llbreak_select_use_line)
    academit.添加命令("llbreak-select-use-select", llbreak_select_use_select)



def llzhu_trans_break_find_select_and_select_acadpointlist(break_objidlist1, refer_objidlist2):
    buflist = []
    for objid1 in break_objidlist1:
        acad_point_list = []
        objref1 = acad.TransObjectForWrite(objid1)
        for objid2 in refer_objidlist2:
            if str(objid1) == str(objid2): continue
            objref2 = acad.TransObjectForWrite(objid2)
            collect = acad.Point3dCollection() 
            objref1.IntersectWith(objref2, acad.Intersect.OnBothOperands, collect, System.IntPtr.Zero, System.IntPtr.Zero)
            for point in collect: 
                acad_point_list.append(point)
        buflist.append([objid1, acad_point_list])
    return buflist

def llzhu_trans_break_objid_with_acadpointlist(objid, acadpointlist):
    objref = acad.TransObjectForWrite(objid)
    pt1, pt2 = [objref.StartPoint.X, objref.StartPoint.Y, objref.StartPoint.Z], [objref.EndPoint.X, objref.EndPoint.Y, objref.EndPoint.Z]
    para_list = []
    for point in acadpointlist:
        # point = objref.GetClosestPointTo(point, extend=False)
        po1 = [point.X, point.Y, point.Z]
        if acad.IsPointSame(pt1, po1): continue
        if acad.IsPointSame(pt2, po1): continue
        para = objref.GetParameterAtPoint(point)
        para_list.append(para)
    para_list.sort() # 默认从小到大
    collect = acad.DoubleCollection()
    for para in para_list:
        collect.Add(para)
    result = objref.GetSplitCurves(collect)
    objref.Erase()
    return result
    

@acad.decorator_command
def llccb():
    height = acad.GetDouble(12, "请输入长城板折高度:")
    width = acad.GetDouble(20, "请输入长城板折宽度:")
    length = acad.GetDouble(1219, "请输入板材长度:")
    pt1 = acad.GetPoint()
    if pt1 == None: return
    ptlist = [pt1]
    sumlength = 0
    while True:
        # 第1条边
        bufheight  = length - sumlength
        sumlength += height
        if abs(sumlength - length) < 0.0001: 
            pt1 = acad.Vec3Add(pt1, [0, -height, 0])
            ptlist.append(pt1)
            break
        if sumlength > length:
            pt1 = acad.Vec3Add(pt1, [0, -bufheight, 0])
            ptlist.append(pt1)
            break
        pt1 = acad.Vec3Add(pt1, [0, -height, 0])
        ptlist.append(pt1)
        # 第2条边
        bufwidth  = length - sumlength
        sumlength += width
        if abs(sumlength - length) < 0.0001: 
            pt1 = acad.Vec3Add(pt1, [width, 0, 0])
            ptlist.append(pt1)
            break
        if sumlength > length:
            pt1 = acad.Vec3Add(pt1, [bufwidth, 0, 0])
            ptlist.append(pt1)
            break
        pt1 = acad.Vec3Add(pt1, [width, 0, 0])
        ptlist.append(pt1)
        # 第3条边
        bufwidth  = length - sumlength
        sumlength += height
        if abs(sumlength - length) < 0.0001: 
            pt1 = acad.Vec3Add(pt1, [0, height, 0])
            ptlist.append(pt1)
            break
        if sumlength > length:
            pt1 = acad.Vec3Add(pt1, [0, bufheight, 0])
            ptlist.append(pt1)
            break
        pt1 = acad.Vec3Add(pt1, [0, height, 0])
        ptlist.append(pt1)
        # 第4条边
        bufwidth  = length - sumlength
        sumlength += width
        if abs(sumlength - length) < 0.0001: 
            pt1 = acad.Vec3Add(pt1, [width, 0, 0])
            ptlist.append(pt1)
            break
        if sumlength > length:
            pt1 = acad.Vec3Add(pt1, [bufwidth, 0, 0])
            ptlist.append(pt1)
            break
        pt1 = acad.Vec3Add(pt1, [width, 0, 0])
        ptlist.append(pt1)


    with acad.transaction() as trans:
        acad.AddLWPolyLine(ptlist)


@acad.decorator_command
def llbreak_point():
    objid = acad.EntSel(string="请选择要打断的对象:")
    ptlist = []
    while True:
        pt1 = acad.GetPoint("请点击打断点:") 
        if pt1 == None: break
        ptlist.append(pt1)
    if ptlist == None: return 

    with acad.transaction() as trans:
        objref = acad.TransObjectForWrite(objid)
        paralist = []
        for pt1 in ptlist:
            po1 = objref.GetClosestPointTo(acad.ToPoint3d(pt1), extend=False)
            pa1 = objref.GetParameterAtPoint(po1)
            paralist.append(pa1)
        if paralist != []:
            paralist.sort()
            collection = acad.DoubleCollection()
            for para in paralist: 
                collection.Add(para) 
            result = objref.GetSplitCurves(collection)
            for i, objrefsub in enumerate(result):
                objrefsub.ColorIndex = i+1
                acad.AddDBObject(objrefsub)
            objref.Erase()


@acad.decorator_command
def llbreak_sublength_for():
    sublength = acad.GetDouble(0, "请输入间隔长度:")
    while True:
        pt0, objid = acad.EntSelEntity(string="请选择要打断的对象:")
        if acad.IsNone(objid): return
        with acad.transaction() as trans:
            objref = acad.TransObjectForWrite(objid)
            count = int(objref.Length/sublength)
            pt1, pt2 = acad.TransEntityStartEndPoint(objid)
            distance1, distance2 = acad.Distance(pt0, pt1), acad.Distance(pt0, pt2)
            if  distance1 <= distance2:
                sumlengthlist = []
                sumlength = 0
                for i in range(count):
                    sumlength += sublength
                    sumlengthlist.append(sumlength)
            else:
                sumlengthlist = []
                sumlength = objref.Length
                for i in range(count):
                    sumlength -= sublength
                    sumlengthlist.append(sumlength)

            paralist = []
            for sumlength in sumlengthlist:
                pa1 = objref.GetParameterAtDistance(sumlength)
                paralist.append(pa1)
            if paralist != []:
                paralist.sort()
                collection = acad.DoubleCollection()
                for para in paralist: 
                    collection.Add(para) 
                result = objref.GetSplitCurves(collection)
                for i, objrefsub in enumerate(result):
                    objrefsub.ColorIndex = i+1
                    acad.AddDBObject(objrefsub)
                objref.Erase()



@acad.decorator_command
def llbreak_subcount_for():
    subcount = acad.GetInt(0, "请输入间隔数量:")
    while True:
        pt0, objid = acad.EntSelEntity(string="请选择要打断的对象:")
        if acad.IsNone(objid): return
        with acad.transaction() as trans:
            objref = acad.TransObjectForWrite(objid)
            sublength = objref.Length/subcount
            sumlengthlist = []
            sumlength = 0
            for i in range(subcount-1):
                sumlength += sublength
                sumlengthlist.append(sumlength)

            paralist = []
            for sumlength in sumlengthlist:
                pa1 = objref.GetParameterAtDistance(sumlength)
                paralist.append(pa1)
            if paralist != []:
                paralist.sort()
                collection = acad.DoubleCollection()
                for para in paralist: 
                    collection.Add(para) 
                result = objref.GetSplitCurves(collection)
                for i, objrefsub in enumerate(result):
                    objrefsub.ColorIndex = i+1
                    acad.AddDBObject(objrefsub)
                objref.Erase()








@acad.decorator_command
def llbreak_select_use_select():# IntersectWith
    break_objidlist = acad.SSGetIdList(string="请选择要打断的对象: ")   
    refer_objidlist = acad.SSGetIdList(string="请选择要参考的对象: ")  
    with acad.transaction() as trans:
        buflist = llzhu_trans_break_find_select_and_select_acadpointlist(break_objidlist, refer_objidlist) 
        for objid1, acad_point_list in buflist:
            result = llzhu_trans_break_objid_with_acadpointlist(objid1, acad_point_list)
            for i, objrefsub in enumerate(result):
                objrefsub.ColorIndex = i
                acad.AddDBObject(objrefsub)


@acad.decorator_command
def llbreak_line_use_select():# IntersectWith # break line or curve
    objid = acad.EntSel(string="请点选要打断的对象:")
    objidlist = acad.SSGetIdList(string="请选择参考的对象:")   
    with acad.transaction() as trans:
        buflist = llzhu_trans_break_find_select_and_select_acadpointlist([objid], objidlist) 
        for objid1, acad_point_list in buflist:
            result = llzhu_trans_break_objid_with_acadpointlist(objid1, acad_point_list)
            for i, objrefsub in enumerate(result):
                objrefsub.ColorIndex = i
                acad.AddDBObject(objrefsub)


@acad.decorator_command
def llbreak_select_use_line():# IntersectWith
    objidlist = acad.SSGetIdList(string="请选择要打断的对象: ")   
    objid = acad.EntSel(string="请点选要参考的对象: ")
    with acad.transaction() as trans:
        buflist = llzhu_trans_break_find_select_and_select_acadpointlist(objidlist, [objid]) 
        for objid1, acad_point_list in buflist:
            result = llzhu_trans_break_objid_with_acadpointlist(objid1, acad_point_list)
            for i, objrefsub in enumerate(result):
                objrefsub.ColorIndex = i
                acad.AddDBObject(objrefsub)


@acad.decorator_command
def llbreak_all():# IntersectWith
    objidlist =  acad.SSGetIdList()   
    with acad.transaction() as trans:
        buflist = llzhu_trans_break_find_select_and_select_acadpointlist(objidlist, objidlist) 
        for objid1, acad_point_list in buflist:
            result = llzhu_trans_break_objid_with_acadpointlist(objid1, acad_point_list)
            for i, objrefsub in enumerate(result):
                objrefsub.ColorIndex = i
                acad.AddDBObject(objrefsub)



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



# GetDistAtParam
# EndParam
# EndPoint
# GetDistAtPoint
# 使用polyline的getSplitCurves方法进行分割曲线时，传入的Point3d点数组，或者Param值数组，必须先进行排序，按从曲线的起点到终点的走向，才能返回正确的曲线段列表。否则切割出来的线段数量不是你想要的。
# 用GetParameterAtPoint将点转为参数，或者GetDistAtPoint将点转为距离，然后按从小到大的顺序进行排列。