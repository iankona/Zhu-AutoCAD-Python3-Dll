import clr

import acad
import academit

import System

def 命令(): 
    academit.添加命令("ll-attach-line-for", ll_attach_line_for)


zhu_llattach_length = 2.0
def zhu_ui_llattach_length():
    global zhu_llattach_length
    length = acad.GetDouble(zhu_llattach_length, string="请输入打标大小: ")
    if length != None: zhu_llattach_length = length


@acad.decorator_command
def ll_attach_line_for():
    zhu_ui_llattach_length()
    while True:
        pt1, pt2 = acad.GetPoint2()
        if pt1 == None: return
        objidlist = acad.GetSelectFenceIdList(pt1, pt2, [[0, "LWPOLYLINE"]])
        with acad.transaction() as trans:
            for objid in objidlist:
                pt1, pt2 = acad.TransStartFinalPoint(objid)   
                po1, po2 = acad.GetAttachNDirectPointList(pt1, pt2, zhu_llattach_length) 
                line1 = acad.AddLine(pt1, po1, layer_name="打标1", color_index=6)
                line2 = acad.AddLine(pt2, po2, layer_name="打标1", color_index=6)

@acad.decorator_command
def ll_attach_rect_length_width_for():
    objidlist = acad.SSGetIdList([[0, "LWPOLYLINE"]])
    with acad.transaction() as trans:
        buflist = []
        for objid in objidlist:
            pt1, pt2 = acad.TransEntityBoundXY0()
            x1, y1, z1 = pt1
            x2, y2, z2 = pt2
            length, width = x2-x1, y2-y1

        # for pt1, pt2 in result:
        #     x1, y1, z1 = pt1
        #     x2, y2, z2 = pt2
        #     length, width = x2-x1, y2-y1
        #     buflist.append([x1, length, width, [pt1, pt2]])
        # buflist.sort(key = lambda item: item[0], reverse=False) # 排序规则，reverse = True 降序， reverse = False 升序（默认） 