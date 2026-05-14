import clr

import acad
import academit

import System

def 命令(): 
    academit.添加命令("llextend-to-circle-for", llextend_to_circle_for)
    academit.添加命令("llextend-to-param-for", llextend_to_param_for)
    academit.添加命令("llextend-to-entity-for", llextend_to_entity_for)
    academit.添加命令("llextend-to-connect-for", llextend_to_connect_for)


@acad.decorator_command
def llextend_to_circle_for():
    r = acad.GetDouble(5, "请输入延长半径:")
    while True:
        pt1, objid1 = acad.EntSelEntity("请选择要延长的对象：")
        if acad.IsNone(objid1): return 
        po1, po2 = acad.GetEntityStartEndPoint(objid1)
        with acad.transaction() as trans:
            objref1 = acad.TransObjectForWrite(objid1)
            if acad.Distance(pt1, po1) < acad.Distance(pt1, po2): # 起点
                circle = acad.DBObjectCircle(po1, r)
                collect = acad.Point3dCollection() 
                objref1.IntersectWith(circle, acad.Intersect.OnBothOperands, collect, System.IntPtr.Zero, System.IntPtr.Zero)
                ptlist1 = [[point.X, point.Y, point.Z] for point in collect]
                collect = acad.Point3dCollection() 
                objref1.IntersectWith(circle, acad.Intersect.ExtendThis, collect, System.IntPtr.Zero, System.IntPtr.Zero)
                ptlist2 = [[point.X, point.Y, point.Z] for point in collect]
                for pt1 in ptlist2:
                    for po1 in ptlist1:
                        if not acad.IsPointSame(pt1, po1): 
                            objref1.Extend(True, acad.ToPoint3d(pt1))
                            break
            else:
                circle = acad.DBObjectCircle(po2, r)
                collect = acad.Point3dCollection() 
                objref1.IntersectWith(circle, acad.Intersect.OnBothOperands, collect, System.IntPtr.Zero, System.IntPtr.Zero)
                ptlist1 = [[point.X, point.Y, point.Z] for point in collect]
                collect = acad.Point3dCollection() 
                objref1.IntersectWith(circle, acad.Intersect.ExtendThis, collect, System.IntPtr.Zero, System.IntPtr.Zero)
                ptlist2 = [[point.X, point.Y, point.Z] for point in collect]
                for pt1 in ptlist2:
                    for po1 in ptlist1:
                        if not acad.IsPointSame(pt1, po1): 
                            objref1.Extend(False, acad.ToPoint3d(pt1))
                            break


@acad.decorator_command
def llextend_to_param_for():
    length = acad.GetDouble(5, "请输入终点延长长度:")
    while True:
        pt1, objid1 = acad.EntSelEntity("请选择要延长的对象：")
        if acad.IsNone(objid1): return 
        po1, po2 = acad.GetEntityStartEndPoint(objid1)
        with acad.transaction() as trans:
            objref1 = acad.TransObjectForWrite(objid1)
            startparam = objref1.StartParam
            endparam = objref1.EndParam
            distance = objref1.GetDistanceAtParameter(endparam)
            dpa = length / distance * (endparam-startparam)
            if acad.Distance(pt1, po1) < acad.Distance(pt1, po2): # 起点
                objref1.Extend(startparam-dpa) # public virtual unsafe void Extend(double newParameter)
            else:
                objref1.Extend(endparam+dpa)



def ptlist1_subtract_ptlist2(ptlist1, ptlist2):
    result = []
    for pt1 in ptlist1: 
        if __is_pt_in_ptlist(pt1, ptlist2): continue
        result.append(pt1)
    return result

def __is_pt_in_ptlist(pt1, ptlist):
    flag = False
    for po1 in ptlist:
        if acad.IsPointSame(pt1, po1): flag = True
    return flag

@acad.decorator_command
def llextend_to_entity_for():
    while True:
        pt1, objid1 = acad.EntSelEntity("请选择要延长的对象：")
        if acad.IsNone(objid1): return 
        pt2, objid2 = acad.EntSelEntity("请选择要抵达的对象：")
        if acad.IsNone(objid2): return 
        po1, po2 = acad.GetEntityStartEndPoint(objid1)
        with acad.transaction() as trans:
            objref1 = acad.TransObjectForWrite(objid1)
            objref2 = acad.TransObjectForWrite(objid2)
            collect = acad.Point3dCollection() 
            objref1.IntersectWith(objref2, acad.Intersect.OnBothOperands, collect, System.IntPtr.Zero, System.IntPtr.Zero)
            ptlist1 = [[point.X, point.Y, point.Z] for point in collect]
            collect = acad.Point3dCollection() 
            objref1.IntersectWith(objref2, acad.Intersect.ExtendThis, collect, System.IntPtr.Zero, System.IntPtr.Zero)
            ptlist2 = [[point.X, point.Y, point.Z] for point in collect]
            ptlist1 = ptlist1_subtract_ptlist2(ptlist2, ptlist1)
            if ptlist1 == []: continue
            pc1 = ptlist1[0]
            if acad.Distance(pc1, po1) < acad.Distance(pc1, po2): # 起点
                objref1.Extend(True, acad.ToPoint3d(pc1)) # public virtual unsafe void Extend([MarshalAs(UnmanagedType.U1)] bool extendStart, Point3d toPoint)
            else:
                objref1.Extend(False, acad.ToPoint3d(pc1))


@acad.decorator_command
def llextend_to_connect_for():
    while True:
        objidlist = acad.SSGetIdList()
        if len(objidlist) < 2 : return
        with acad.transaction() as trans:
            objref1 = acad.TransObjectForWrite(objidlist[0])
            objref2 = acad.TransObjectForWrite(objidlist[1])
            collect = acad.Point3dCollection() 
            objref1.IntersectWith(objref2, acad.Intersect.OnBothOperands, collect, System.IntPtr.Zero, System.IntPtr.Zero)
            ptlist1 = [[point.X, point.Y, point.Z] for point in collect]
            collect = acad.Point3dCollection() 
            objref1.IntersectWith(objref2, acad.Intersect.ExtendBoth, collect, System.IntPtr.Zero, System.IntPtr.Zero)
            ptlist2 = [[point.X, point.Y, point.Z] for point in collect]
            ptlist1 = ptlist1_subtract_ptlist2(ptlist2, ptlist1)
            if ptlist1 == []: continue
            pc1 = ptlist1[0]

            po1, po2 = acad.GetEntityStartEndPoint(objidlist[0])
            if acad.Distance(pc1, po1) < acad.Distance(pc1, po2): # 起点
                objref1.Extend(True, acad.ToPoint3d(pc1)) # public virtual unsafe void Extend([MarshalAs(UnmanagedType.U1)] bool extendStart, Point3d toPoint)
            else:
                objref1.Extend(False, acad.ToPoint3d(pc1))

            po1, po2 = acad.GetEntityStartEndPoint(objidlist[1])
            if acad.Distance(pc1, po1) < acad.Distance(pc1, po2): # 起点
                objref2.Extend(True, acad.ToPoint3d(pc1)) # public virtual unsafe void Extend([MarshalAs(UnmanagedType.U1)] bool extendStart, Point3d toPoint)
            else:
                objref2.Extend(False, acad.ToPoint3d(pc1))


# ExtendBoth 两个实体都延伸
# ExtendArgument 只延伸作为参数的实体(该方法的第一个参数)
# ExtendThis 只延伸原实体（调用该方法的实体）
# OnBothOperands 两个实体都不延伸