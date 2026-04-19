import clr

import acad
import academit

import System

def 命令(): 
    academit.添加命令("llsc-same-for", llsc_same_for)
    academit.添加命令("llsc-diff-for", llsc_diff_for)

@acad.decorator_command_undo
def llsc_same_for():
    length = acad.GetDouble(1000, "请输入水槽总长度:")
    while True:
        pt1, objid1 = acad.EntSelEntity("请选择第1条多段线:")
        pa1 = acad.GetPoint()
        if pa1 == None: return
        ps1, pe1 = acad.GetEntityStartEndPoint(objid1)
        lengthlist1 = acad.GetLWPolyLineLengthList(objid1)
        if acad.Distance(pt1, ps1) > acad.Distance(pt1, pe1): 
            lengthlist1 = lengthlist1[::-1]
        with acad.transaction() as trans:
            pt1 = pa1
            po1 = acad.Vec3Add(pt1, [length, 0])
            acad.AddLine(pt1, po1)
            for length1 in lengthlist1:
                pt2 = acad.Vec3Add(pt1, [0, -length1])
                po2 = acad.Vec3Add(po1, [0, -length1])
                line1 = acad.AddLine(pt1, pt2)
                line2 = acad.AddLine(po1, po2)
                line3 = acad.AddLine(pt2, po2, "图层1")
                # pa1, pa2 = acad.GetAttachNDirectPt1Pt2(pt2, po2)
                # line4 = acad.AddLine(pt2, pa1)
                # line5 = acad.AddLine(po2, pa2)            
                pt1, po1 = pt2, po2
            line3.Layer = "0"
            # line4.Erase()
            # line5.Erase()


@acad.decorator_command_undo
def llsc_diff_for():
    length = acad.GetDouble(1000, "请输入水槽总长度:")
    while True:
        pt1, objid1 = acad.EntSelEntity("请选择第1条多段线:")
        pt2, objid2 = acad.EntSelEntity("请选择第2条多段线:")
        pd1, pd2 = acad.GetPoint2("请选择对齐点1:", "请选择对齐点2:")
        if pd1 == None: return 
        pa1 = acad.GetPoint()
        if pa1 == None: return 
        ps1, pe1 = acad.GetEntityStartEndPoint(objid1)
        ps2, pe2 = acad.GetEntityStartEndPoint(objid2)
        sumlength1, offlength1 = acad.GetLWPolyLineLengthAndLengthAtPoint(objid1, pd1) 
        sumlength2, offlength2 = acad.GetLWPolyLineLengthAndLengthAtPoint(objid2, pd2) 
        lengthlist1 = acad.GetLWPolyLineLengthList(objid1)
        lengthlist2 = acad.GetLWPolyLineLengthList(objid2)
        if acad.Distance(pt1, ps1) > acad.Distance(pt1, pe1): 
            lengthlist1 = lengthlist1[::-1]
            offlength1 = sumlength1 - offlength1
        if acad.Distance(pt2, ps2) > acad.Distance(pt2, pe2): 
            lengthlist2 = lengthlist2[::-1]
            offlength2 = sumlength2 - offlength2
        dy = offlength2 - offlength1

        with acad.transaction() as trans:
            pt1 = pa1
            po1 = acad.Vec3Add(pt1, [length, dy])
            acad.AddLine(pt1, po1)
            for length1, length2 in zip(lengthlist1,lengthlist2):
                pt2 = acad.Vec3Add(pt1, [0, -length1])
                po2 = acad.Vec3Add(po1, [0, -length2])
                line1 = acad.AddLine(pt1, pt2)
                line2 = acad.AddLine(po1, po2)
                line3 = acad.AddLine(pt2, po2, "图层1")
                # pa1, pa2 = acad.GetAttachNDirectPt1Pt2(pt2, po2)
                # line4 = acad.AddLine(pt2, pa1)
                # line5 = acad.AddLine(po2, pa2)            
                pt1, po1 = pt2, po2
            line3.Layer = "0"
            # line4.Erase()
            # line5.Erase()




    



