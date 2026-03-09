import clr

import acad
import academit

import System

def 命令(): 
    academit.添加命令("llcr-circle3d", llcr_circle3d)
    academit.添加命令("llcd-circle3d", llcd_circle3d)



zhu_circle_r = 50
@acad.decorator_command
def llcr_circle3d():
    global zhu_circle_r
    r = acad.GetDouble(zhu_circle_r, "请输入圆的半径:")
    if r != None: zhu_circle_r = r

    [pickpoint, objid] = acad.EntSelEntity()
    startpoint, endpoint = acad.GetStartFinalPoint(objid)
    dis_s, dis_e = acad.Distance(pickpoint, startpoint), acad.Distance(pickpoint, endpoint)
    if dis_s < dis_e:
        center, normal = startpoint, acad.Direct(startpoint, endpoint)
    else:
        center, normal = endpoint, acad.Direct(endpoint, startpoint)
    with acad.transaction() as trans:
        acad.AddCircle(center, r, normal)

    

@acad.decorator_command
def llcd_circle3d():
    global zhu_circle_r
    d = acad.GetDouble(zhu_circle_r*2, "请输入圆的直径:")
    if d != None: zhu_circle_r = d/2

    [pickpoint, objid] = acad.EntSelEntity()
    startpoint, endpoint = acad.GetStartFinalPoint(objid)
    dis_s, dis_e = acad.Distance(pickpoint, startpoint), acad.Distance(pickpoint, endpoint)
    if dis_s < dis_e:
        center, normal = startpoint, acad.Direct(startpoint, endpoint)
    else:
        center, normal = endpoint, acad.Direct(endpoint, startpoint)
    with acad.transaction() as trans:
        acad.AddCircle(center, r, normal)