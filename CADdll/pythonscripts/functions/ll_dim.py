

import acad
import academit
import math
import math
import numpy as np
import time 
def 命令(): 
    academit.添加命令("lldim-for", lldim_for)
    academit.添加命令("lldal3d", lldal3d)

# Application.DocumentManager.MdiActiveDocument.Editor.PointMonitor

zhu_dim_biaohao = 15

@acad.decorator_command
def lldim_for():
    global zhu_dim_biaohao
    biaohao = acad.GetInt(zhu_dim_biaohao, "请输入标注号:")
    if biaohao != None: zhu_dim_biaohao = biaohao
    while True:
        pt1, pt2, pt3 = acad.GetPoint3("请选择第1个顶点:", "请选择第2个顶点:", "请点击方向顶点:")
        if pt1 == None: break
        with acad.transaction() as trans:
            acad.AddDim(pt1, pt2, pt3)



zhu_dim_height = 250
@acad.decorator_command
def lldal3d():
    global zhu_dim_height
    height = acad.GetDouble(zhu_dim_height, "请输入标注高度:")
    if height != None: zhu_dim_height = height        
    pt1, pt2, pt3 = acad.GetPoint3("请选择标注第1个顶点:", "请选择标注第2个顶点:", "请选择标注第3个顶点:")
    normal = acad.Cross(acad.Direct(pt1, pt2), acad.Direct(pt2, pt3))
    length = acad.Distance(pt1, pt2)
    po1, po2 = pt1[0:2]+[0], pt2[0:2]+[0]
    dr1 = acad.Direct(po1, po2)
    dr1 = acad.Vec3ResetLength(dr1, length)
    po2 = acad.Vec3Add(po1, dr1)
    mid = acad.MidPt1Pt2(po1, po2)
    dr2 = acad.Vec3ResetLength(normal, height)
    po3 = acad.Vec3Add(mid, dr2)

    rotation_angle = acad.AngleFromDotDr1Dr2(acad.Direct(pt1, pt2), acad.Direct(po1, po2))
    with acad.transaction() as trans:
        dim = acad.AddDal(po1, po2, po3)
        acad.TransMove(dim.ObjectId, po1, pt1)
        acad.TransRoation(dim.ObjectId, rotation_angle, normal, pt1)



def Matrix3x3Dr1ToDr2(dr1, dr2):
    dr2 = np.array([1, 0 ,0])
    dr1 = np.array([0, 0 ,1])
    #注意，如果向量没有归一化，可以先考虑归一化下。
    c = np.dot(dr2, dr1)
    normal = np.cross(dr2, dr1)
    s = np.linalg.norm(normal)
    print(c,s)
    normal_invert = np.array((
    [         0, -normal[2],  normal[1]],
    [ normal[2],          0, -normal[0]],
    [-normal[1],  normal[0],          0],
    ))
    I = np.eye(3)
    # 核心公式:见上图
    R_w2c = I + normal_invert + np.dot(normal_invert, normal_invert)/(1+c)
    print(R_w2c)