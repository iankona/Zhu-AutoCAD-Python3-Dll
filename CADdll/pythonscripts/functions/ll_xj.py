import clr

import acad
import academit

import System
import math


def 命令(): 
    academit.添加命令("llxj-for", llxj_for)
    academit.添加命令("llxj-w-for", llxj_w_for)
    academit.添加命令("llxj-angle-for", llxj_angle_for)


@acad.decorator_command
def llxj_for():
    while True:
        pt1, pt2 = acad.GetPoint2()
        if pt1 == None: return
        dr1 = acad.GetPerDirectWithPerflagXY(pt1, pt2,  1)
        dr2 = acad.GetPerDirectWithPerflagXY(pt1, pt2, -1)
        po1 = acad.Vec3Add(pt2, dr1)
        po2 = acad.Vec3Add(pt2, dr2)
        with acad.transaction() as trans:
            acad.AddLine(pt1, po1)
            acad.AddLine(pt1, po2)

llzhu_xj_ness = 1.0
llzhu_xj_ness_half = 0.5
def uizhu_xj_ness():
    global llzhu_xj_ness, llzhu_xj_ness_half
    ness = acad.GetDouble(llzhu_xj_ness, string="请输入板厚:")
    if ness != None: 
        llzhu_xj_ness = ness
        llzhu_xj_ness_half = ness / 2


@acad.decorator_command
def llxj_w_for():
    uizhu_xj_ness()
    while True:
        pt1, pt2 = acad.GetPoint2()
        if pt1 == None: return
        dr0 = acad.Direct(pt1, pt2)
        dr1 = acad.GetPerDirectWithPerflagXY(pt1, pt2,  1)
        dr2 = acad.GetPerDirectWithPerflagXY(pt1, pt2, -1)
        dr3 = acad.Vec3ResetLength(dr1, llzhu_xj_ness)
        dr4 = acad.Vec3ResetLength(dr2, llzhu_xj_ness)
        dr1 = acad.Vec3Add(dr0, dr1)
        dr2 = acad.Vec3Add(dr0, dr2)
        with acad.transaction() as trans:
            po1 = acad.Vec3Add(pt1, dr3)
            po2 = acad.Vec3Add(pt1, dr4)
            acad.AddLine(pt1, po1)
            acad.AddLine(pt1, po2)
            pd1 = acad.Vec3Add(po1, dr1)
            pd2 = acad.Vec3Add(po2, dr2)
            acad.AddLine(po1, pd1)
            acad.AddLine(po2, pd2)


llzhu_xj_angle = 30
llzhu_xj_angle_half = 15
@acad.decorator_command
def llxj_angle_for():
    global llzhu_xj_angle
    angle = acad.GetDouble(llzhu_xj_angle, string="请输入两条斜边夹成的角度:")
    if angle != None: 
        llzhu_xj_angle = angle
        llzhu_xj_angle_half = angle/2
    while True:
        pt1, pt2 = acad.GetPoint2()
        if pt1 == None: return
        length = acad.Distance(pt1, pt2)
        dr1 = acad.GetPerDirectWithPerflagResetLengthXY(pt1, pt2,  1, length*math.tan(acad.Angle2Rad(llzhu_xj_angle_half)))
        dr2 = acad.GetPerDirectWithPerflagResetLengthXY(pt1, pt2, -1, length*math.tan(acad.Angle2Rad(llzhu_xj_angle_half)))
        po1 = acad.Vec3Add(pt2, dr1)
        po2 = acad.Vec3Add(pt2, dr2)
        with acad.transaction() as trans:
            acad.AddLine(pt1, po1)
            acad.AddLine(pt1, po2)
    