  
import acad
import academit
import System

def 命令(): 
    academit.添加命令("ll-break-point", ll_break_point)
    academit.添加命令("ll-break-line-use-select", ll_break_line_use_select)
    academit.添加命令("ll-break-select-use-line", ll_break_select_use_line)
    academit.添加命令("ll-break-all", ll_break_all)


@acad.decorator_command
def ll_break_point():
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
def ll_break_line_use_select():# IntersectWith # break line or curve
    objid = acad.EntSel(string="请选择要打断的对象:")
    objidlist = acad.SSGetIdList(string="请选择参考线:")   
    with acad.transaction() as trans:
        point_list = []
        objref1 = acad.TransObjectForWrite(objid)
        for objid2 in objidlist:
            if str(objid) == str(objid2): continue
            collect = acad.Point3dCollection() 
            objref2 = acad.TransObjectForWrite(objid2)
            objref1.IntersectWith(objref2, acad.Intersect.OnBothOperands, collect, System.IntPtr.Zero, System.IntPtr.Zero)
            for point in collect: 
                point_list.append(point)

        if point_list != []:
            para_list = []
            for point in point_list:
                para = objref1.GetParameterAtPoint(point)
                para_list.append(para)
            para_list.sort()
            collect = acad.DoubleCollection()
            for para in para_list:
                collect.Add(para)
            result = objref1.GetSplitCurves(collect)
            for i, objrefsub in enumerate(result):
                objrefsub.ColorIndex = i
                acad.AddDBObject(objrefsub)
            objref1.Erase()



@acad.decorator_command
def ll_break_select_use_line():# IntersectWith
    objidlist = acad.SSGetIdList(string="请选择要打断的对象: ")   
    objid2 = acad.EntSel(string="请选择参考线: ")
    with acad.transaction() as trans:
        objref2 = acad.TransObjectForWrite(objid2)
        for objid1 in objidlist:
            objref1 = acad.TransObjectForWrite(objid1)
            collect = acad.Point3dCollection() 
            objref1.IntersectWith(objref2, acad.Intersect.OnBothOperands, collect, System.IntPtr.Zero, System.IntPtr.Zero)
            point_list = []
            for point in collect: 
                point_list.append(point)

            if point_list != []:
                para_list = []
                for point in point_list:
                    para = objref1.GetParameterAtPoint(point)
                    para_list.append(para)
                para_list.sort()
                collect = acad.DoubleCollection()
                for para in para_list:
                    collect.Add(para)
                result = objref1.GetSplitCurves(collect)
                for i, objrefsub in enumerate(result):
                    objrefsub.ColorIndex = i
                    acad.AddDBObject(objrefsub)
                objref1.Erase()


@acad.decorator_command
def ll_break_all():# IntersectWith
    objidlist =  acad.SSGetIdList()   
    count = len(objidlist)
    objid_point_list = []
    with acad.transaction() as trans:
        for i in range(count):
            objid1 = objidlist[i]
            objref1 = acad.TransObjectForWrite(objid1)
            for objid2 in objidlist[i+1:]:
                collect = acad.Point3dCollection() 
                objref2 = acad.TransObjectForWrite(objid2)
                objref1.IntersectWith(objref2, acad.Intersect.OnBothOperands, collect, System.IntPtr.Zero, System.IntPtr.Zero)
                for point in collect:
                    objid_point_list.append([str(objid1), point])
                    objid_point_list.append([str(objid2), point])
        count = 1
        for objid in objidlist:
            point_list = []
            for objidstring, point in objid_point_list:
                if str(objid) == objidstring: point_list.append(point)
            if point_list == []: continue
            objref = acad.TransObjectForWrite(objid)
            para_list = []
            for point in point_list:
                para = objref.GetParameterAtPoint(point)
                para_list.append(para)
            para_list.sort()
            collect = acad.DoubleCollection()
            for para in para_list:
                collect.Add(para)
            result = objref.GetSplitCurves(collect)
            for objrefsub in result:
                objrefsub.ColorIndex = count
                acad.AddDBObject(objrefsub)
                count += 1
            objref.Erase()


# GetDistAtParam
# EndParam
# EndPoint
# GetDistAtPoint
# 使用polyline的getSplitCurves方法进行分割曲线时，传入的Point3d点数组，或者Param值数组，必须先进行排序，按从曲线的起点到终点的走向，才能返回正确的曲线段列表。否则切割出来的线段数量不是你想要的。
# 用GetParameterAtPoint将点转为参数，或者GetDistAtPoint将点转为距离，然后按从小到大的顺序进行排列。