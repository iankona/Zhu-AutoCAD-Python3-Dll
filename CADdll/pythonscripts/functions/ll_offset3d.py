


import clr

import acad
import academit

import System

def 命令(): 
    # academit.添加命令("llyx", llyx)
    # academit.添加命令("lljz_gaiban_njfor", lljz_gaiban_njfor) # 盖板
    # academit.添加命令("lljz_geban_njfor", lljz_geban_njfor) # 隔板
    pass

tube_length = 40

@acad.decorator_command
def lljz_gaiban_njfor(): # 盖板
    global tube_length
    length = acad.GetDouble(tube_length, f"请输入方管宽度: ") 
    if length != None: tube_length = length
    length_list = [tube_length, tube_length, 15]
    objidlist = acad.SSGetIdList([[0, "LWPOLYLINE"]])
    with acad.transaction() as trans:
        acad.ChangeObjectIdLayer(objidlist, "图层1")
        for objid in objidlist:
            center = acad.GetEntityBoundCenterXY(objid)
            pline_point_list = acad.GetLWPolyLinePointList(objid)
            for i in range(len(pline_point_list)-1):
                pt1 = pline_point_list[i]
                pt2 = pline_point_list[i+1]
                pt3 = center
                perflag = acad.GetPerflagXY(pt1, pt2, pt3)
                perflag = -perflag
                dr1 = acad.GetPerDirectWithPerflagXY(pt1, pt2, perflag) 
                for m, length in enumerate(length_list):
                    dr1 = acad.Vec3ResetLength(dr1, length)
                    match m:
                        case 0:
                            po1 = acad.Vec3Add(pt1, dr1)
                            po2 = acad.Vec3Add(pt2, dr1)
                            line1 = acad.AddLine(pt1, po1)
                            line2 = acad.AddLine(pt2, po2)
                        case 1:
                            pn1, pn2 = acad.GetAttachNDirectPointList(pt1, pt2, length)
                            line1 = acad.AddLine(pt1, pn1)
                            line2 = acad.AddLine(pt2, pn2)
                            line3 = acad.AddLine(pn1, pn2, "图层1")
                            po1 = acad.Vec3Add(pn1, dr1)
                            po2 = acad.Vec3Add(pn2, dr1)
                            line1 = acad.AddLine(pn1, po1)
                            line2 = acad.AddLine(pn2, po2)
                            line3 = acad.AddLine(po1, po2, "图层1")                            
                        case _:
                            po1 = acad.Vec3Add(pt1, dr1)
                            po2 = acad.Vec3Add(pt2, dr1)
                            line1 = acad.AddLine(pt1, po1)
                            line2 = acad.AddLine(pt2, po2)
                            line3 = acad.AddLine(po1, po2, "图层1")
                    pt1, pt2 = po1, po2

                line3.Layer = "0"


# 
