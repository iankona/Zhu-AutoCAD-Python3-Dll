import clr

import acad
import academit

import System

def 命令(): 
    academit.添加命令("llpl-to-midpl", llpl_to_midpl)
    pass


@acad.decorator_command
def llpl_to_midpl():
    objidlist = acad.SSGetIdList([[0, "LWPOLYLINE"]])
    with acad.command_undo(), acad.osmode():
        for objid in objidlist:
            ptlist = []
            pline_point_list = acad.GetLWPolyLinePointList(objid)
            ptlist.append(pline_point_list[0])
            for i in range(len(pline_point_list)-1):
                pt1 = pline_point_list[i]
                pt2 = pline_point_list[i+1]
                ptlist.append(acad.MidPt1Pt2(pt1, pt2))
            ptlist.append(pline_point_list[-1])
            acad.CommandAddPLine(ptlist)

