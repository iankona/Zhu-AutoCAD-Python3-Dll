
import acad
import academit

import open3d as o3d
import numpy as np

def 命令(): 
    academit.添加命令("llrect-point-select-add-square-w-for", llrect_point_select_add_square_w_for)
    pass



zhu_square_size = 25
@acad.decorator_command
def llrect_point_select_add_square_w_for():
    import ll_connet
    global zhu_square_size
    length = acad.GetDouble(zhu_square_size, "请输入外方管大小:")
    if length != None: zhu_square_size = length

    while True:
        objid1 = acad.EntSel([[0, "LWPOLYLINE"]], string="请点击第1条对象:") 
        if acad.IsNone(objid1): return

        with acad.transaction() as trans:
            objref1 = acad.TransObjectForRead(objid1)
            ptnlist = []
            for i in range(objref1.NumberOfVertices):
                point = objref1.GetPoint3dAt(i)
                ptnlist.append([point.X, point.Y, point.Z])
            if acad.IsCCW(ptnlist):
                collect = objref1.GetOffsetCurves( zhu_square_size)  
            else:
                collect = objref1.GetOffsetCurves(-zhu_square_size) 
            objref2 = collect[0] 
            ptwlist = []
            for i in range(objref2.NumberOfVertices):
                point = objref2.GetPoint3dAt(i)
                ptwlist.append([point.X, point.Y, point.Z])
            buflist = ll_connet.zhu_connet_rect_and_rect(ptnlist, ptwlist)
            for pt1, pt2 in buflist:
                x1, y1, z1 = pt1
                x2, y2, z2 = pt2
                po1 = [x1, y1, 0]
                po2 = [x1, y2, 0]
                po3 = [x2, y2, 0]
                po4 = [x2, y1, 0]
                acad.AddLWPolyLine([po1, po2, po3, po4, po1])




# objidlist = acad.SSGetIdList([[0, "LWPOLYLINE"]])