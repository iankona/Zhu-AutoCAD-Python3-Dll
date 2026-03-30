

import acad
import academit
import math
import math
import numpy as np
import time 
def 命令(): 
    academit.添加命令("lldim-for", lldim_for)
    academit.添加命令("lldim-entsel-for", lldim_entsel_for)
    academit.添加命令("lldim-fence-qdim-x-for", lldim_fence_qdim_x_for)
    academit.添加命令("lldim-fence-qdim-y-for", lldim_fence_qdim_y_for)



llzhu_dim_biaohao = 15
llzhu_dim_height = 15
def llzhu_ui_dim_input():
    global llzhu_dim_biaohao, llzhu_dim_height
    biaohao = acad.GetInt(llzhu_dim_biaohao, "请输入标注号:")
    height = acad.GetInt(llzhu_dim_height, "请输入标注高:")
    if biaohao != None: llzhu_dim_biaohao = biaohao
    if height != None: llzhu_dim_height = height

llzhu_dim_direct_flag = 1
def llzhu_ui_dim_direct_flag():
    global llzhu_dim_direct_flag
    flag = acad.GetInt(llzhu_dim_direct_flag, "请输入标注方向+X+Y或-X-Y方向:")
    if flag > 0: llzhu_dim_direct_flag = 1
    else: llzhu_dim_direct_flag = -1



@acad.decorator_command
def lldim_for():
    llzhu_ui_dim_input()
    while True:
        pt1, pt2, pt3 = acad.GetPoint3("请选择第1个顶点:", "请选择第2个顶点:", "请点击方向顶点:")
        if pt1 == None: break
        with acad.transaction() as trans:
            acad.AddDim(pt1, pt2, pt3)



llzhu_dim_filter = [
    [-4, "<OR"], [0, "LINE"], [0, "LWPOLYLINE"], [-4, "OR>"]] # , [0, "SPLINE"],  [0, "ARC"], [0, "CIRCLE"], [0, "ELLIPSE"], [-4, "OR>"]]

@acad.decorator_command
def lldim_entsel_for():
    llzhu_ui_dim_input()
    while True:
        pd1, objid = acad.EntSelEntity()
        if acad.IsNone(objid): return
        pt3 = acad.GetPoint()
        if pt3 == None: return
        with acad.transaction() as trans:
            pt1, pt2 = acad.TransEntityStartEndPoint(objid)
            dr1 = acad.GetPerDirectResetLengthXY(pt1, pt2, pt3, llzhu_dim_height)
            pd1 = acad.Vec3Add(pd1, dr1)
            acad.TransCurrentDimStyle(f"副本{llzhu_dim_biaohao} ISO-25")
            acad.AddDal(pt1, pt2, pd1)



@acad.decorator_command
def lldim_fence_qdim_x_for():
    llzhu_ui_dim_input()
    llzhu_ui_dim_direct_flag()
    while True:
        pt1, pt2 = acad.GetPoint2()
        if pt1 == None: break
        objidlist = acad.GetSelectFenceIdList(pt1, pt2, llzhu_dim_filter)
        [x1, y1, z1], [x2, y2, z2] = acad.GetIdListBoundXY0(objidlist)

        if llzhu_dim_direct_flag > 0:
            pd3 = [(x1+x2)/2, y2+llzhu_dim_height, 0]
        else:
            pd3 = [(x1+x2)/2, y1-llzhu_dim_height, 0]
        ss1 = acad.SSSetFromIdList(objidlist)
        acad.SetCurrentDimStyle(f"副本{llzhu_dim_biaohao} ISO-25")
        acad.SetEntNext()
        acad.Command(["qdim", ss1, "", acad.ToPoint3d(pd3)])
        objidlist = acad.GetEntNextIdList()
        # with acad.transaction() as trans:
        #     result = acad.TransAutoMoveXDimTextPointList(objidlist)
        #     for objid, pt0 in result:
        #         objref = acad.TransObjectForWrite(objid)
        #         objref.TextPosition = acad.ToPoint3d(pt0)
                



@acad.decorator_command
def lldim_fence_qdim_y_for():
    llzhu_ui_dim_input()
    llzhu_ui_dim_direct_flag()
    while True:
        pt1, pt2 = acad.GetPoint2()
        if pt1 == None: break
        objidlist = acad.GetSelectFenceIdList(pt1, pt2, llzhu_dim_filter)
        [x1, y1, z1], [x2, y2, z2] = acad.GetIdListBoundXY0(objidlist)
        if llzhu_dim_direct_flag > 0:
            pd3 = [x2+llzhu_dim_height, (y1+y2)/2, 0]
        else:
            pd3 = [x1-llzhu_dim_height, (y1+y2)/2, 0]
        ss1 = acad.SSSetFromIdList(objidlist)
        acad.SetCurrentDimStyle(f"副本{llzhu_dim_biaohao} ISO-25")
        acad.Command(["qdim", ss1, "", acad.ToPoint3d(pd3)])



# with acad.transaction() as trans:
#     buflist = []
#     for objjid in objidlist: 
#         objref = acad.TransObjectForRead(objjid)
#         if objref.Layer == "打标1": continue
#         buflist.append(objjid)

# zhu_dim_height = 250
# @acad.decorator_command
# def lldal3d():
#     global zhu_dim_height
#     height = acad.GetDouble(zhu_dim_height, "请输入标注高度:")
#     if height != None: zhu_dim_height = height        
#     pt1, pt2, pt3 = acad.GetPoint3("请选择标注第1个顶点:", "请选择标注第2个顶点:", "请选择标注第3个顶点:")
#     normal = acad.Cross(acad.Direct(pt1, pt2), acad.Direct(pt2, pt3))
#     length = acad.Distance(pt1, pt2)
#     po1, po2 = pt1[0:2]+[0], pt2[0:2]+[0]
#     dr1 = acad.Direct(po1, po2)
#     dr1 = acad.Vec3ResetLength(dr1, length)
#     po2 = acad.Vec3Add(po1, dr1)
#     mid = acad.MidPt1Pt2(po1, po2)
#     dr2 = acad.Vec3ResetLength(normal, height)
#     po3 = acad.Vec3Add(mid, dr2)

#     rotation_angle = acad.AngleFromDotDr1Dr2(acad.Direct(pt1, pt2), acad.Direct(po1, po2))
#     with acad.transaction() as trans:
#         dim = acad.AddDal(po1, po2, po3)
#         acad.TransMove(dim.ObjectId, po1, pt1)
#         acad.TransRoation(dim.ObjectId, rotation_angle, normal, pt1)



# def Matrix3x3Dr1ToDr2(dr1, dr2):
#     dr2 = np.array([1, 0 ,0])
#     dr1 = np.array([0, 0 ,1])
#     #注意，如果向量没有归一化，可以先考虑归一化下。
#     c = np.dot(dr2, dr1)
#     normal = np.cross(dr2, dr1)
#     s = np.linalg.norm(normal)
#     print(c,s)
#     normal_invert = np.array((
#     [         0, -normal[2],  normal[1]],
#     [ normal[2],          0, -normal[0]],
#     [-normal[1],  normal[0],          0],
#     ))
#     I = np.eye(3)
#     # 核心公式:见上图
#     R_w2c = I + normal_invert + np.dot(normal_invert, normal_invert)/(1+c)
#     print(R_w2c)