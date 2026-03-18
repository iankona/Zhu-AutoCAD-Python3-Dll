  
import acad
import academit
import System

def 命令(): 
    academit.添加命令("llbreak-all", llbreak_all)
    academit.添加命令("llbreak-point", llbreak_point)
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




# GetDistAtParam
# EndParam
# EndPoint
# GetDistAtPoint
# 使用polyline的getSplitCurves方法进行分割曲线时，传入的Point3d点数组，或者Param值数组，必须先进行排序，按从曲线的起点到终点的走向，才能返回正确的曲线段列表。否则切割出来的线段数量不是你想要的。
# 用GetParameterAtPoint将点转为参数，或者GetDistAtPoint将点转为距离，然后按从小到大的顺序进行排列。