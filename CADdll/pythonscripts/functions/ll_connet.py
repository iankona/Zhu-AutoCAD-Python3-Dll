import clr

import acad
import academit

import System

def 命令(): 
    academit.添加命令("llconnet-line-and-line-for", llconnet_line_and_line_for)
    academit.添加命令("llconnet-pl-and-pl-for", llconnet_pl_and_pl_for)
    # academit.添加命令("llconnet-rect-and-rect-simple", llconnet_rect_and_rect_simple)
    # academit.添加命令("llconnet-rect-and-rect-point", llconnet_rect_and_rect_point)
    academit.添加命令("llconnet-rect-and-rect-point-for", llconnet_rect_and_rect_point_for)
    academit.添加命令("llconnet-rect-and-rect-perx-for", llconnet_rect_and_rect_perx_for)
    academit.添加命令("llconnet-rect-and-rect-pery-for", llconnet_rect_and_rect_pery_for)
    pass



def zhu_connet_rect_and_rect(ptnlist, ptwlist):
    buflist = []
    for pt1 in ptnlist:
        lengthlist = []
        for pt2 in ptwlist:
            length = acad.Distance(pt1, pt2)
            lengthlist.append([pt2, length])
        lengthlist.sort(key = lambda item: item[1], reverse=False) # 排序规则，reverse = True 降序， reverse = False 升序（默认）
        pt2 = lengthlist[0][0]
        buflist.append([pt1, pt2])
    return buflist


def llzhu_trans_connect_line_and_line(objidlist):
    result = []
    objid1, objid2 = objidlist[0:2]
    pt1, mid1, pt2 = acad.TransEntityStartMidEndPoint(objid1)
    po1, mid2, po2 = acad.TransEntityStartMidEndPoint(objid2)
    dr1 = acad.Direct(mid2, mid1)
    dr1 = acad.Vec3ResetLength(dr1, 18)
    pt4 = acad.Vec3Add(pt1, dr1)
    pt3 = acad.Vec3Add(pt2, dr1)
    result.append([pt2, pt3])
    result.append([pt3, pt4])
    result.append([pt1, pt4])
    dr1 = acad.Direct(mid1, mid2)
    dr1 = acad.Vec3ResetLength(dr1, 18)
    po4 = acad.Vec3Add(po1, dr1)
    po3 = acad.Vec3Add(po2, dr1)
    result.append([po2, po3])
    result.append([po3, po4])
    result.append([po1, po4])
    pa1, pa2 = acad.GetAttachNDirectPointList(pt1, pt2, 2.0)
    result.append([pt1, pa1])
    result.append([pt2, pa2])
    pa1, pa2 = acad.GetAttachNDirectPointList(po1, po2, 2.0)
    result.append([po1, pa1])
    result.append([po2, pa2])
    length1 = acad.Distance(pt1, po1)
    length2 = acad.Distance(pt1, po2)
    if length1 < length2: 
        result.append([pt1, po1])
    else:
        result.append([pt1, po2])
    length1 = acad.Distance(pt2, po1)
    length2 = acad.Distance(pt2, po2)
    if length1 < length2: 
        result.append([pt2, po1])
    else:
        result.append([pt2, po2])
    return result


# @acad.decorator_command
# def llconnet_line_and_line_set_for():
#     while True: 
#         pt1, pt2 = acad.GetPoint2()
#         if pt1 == None: return
#         objidlist = acad.GetSelectFenceIdList(pt1, pt2)
#         with acad.transaction() as trans:
#             result = llzhu_trans_connect_line_and_line(objidlist)
#             for pt1, pt2 in result: 
#                 if acad.Distance(pt1, pt2) < 2.1:
#                     acad.AddLine(pt1, pt2, "打标1", 6)
#                 else:
#                     acad.AddLine(pt1, pt2)
        


@acad.decorator_command
def llconnet_line_and_line_for():
    while True: 
        pt1, pt2 = acad.GetPoint2()
        if pt1 == None: return
        objidlist = acad.GetSelectFenceIdList(pt1, pt2)
        pt1, pt2 = acad.GetEntityStartEndPoint(objidlist[0])
        po1, po2 = acad.GetEntityStartEndPoint(objidlist[1])
        if acad.Distance(pt1, po1) > acad.Distance(pt1, po2): po1, po2 = po2, po1
        with acad.transaction() as trans:
            acad.AddLine(pt1, po1)
            acad.AddLine(pt2, po2)


@acad.decorator_command
def llconnet_pl_and_pl_for():
    while True: 
        pt1, pt2 = acad.GetPoint2()
        if pt1 == None: return
        objidlist = acad.GetSelectFenceIdList(pt1, pt2)
        ptlist1 = acad.GetLWPolyLinePointList(objidlist[0])
        ptlist2 = acad.GetLWPolyLinePointList(objidlist[1])
        pt1, pt2 = ptlist1[0], ptlist1[-1]
        po1, po2 = ptlist2[0], ptlist2[-1]
        if acad.Distance(pt1, po1) > acad.Distance(pt1, po2): ptlist2 = ptlist2[::-1]
        with acad.transaction() as trans:
            for pt1, po1 in zip(ptlist1, ptlist2):
                acad.AddLine(pt1, po1)



# @acad.decorator_command
# def llconnet_pl_and_pl_set_for():
#     while True:
#         objid1 = acad.EntSel(string="请点击第1条对象:") # [[0, "LWPOLYLINE"]]
#         if acad.IsNone(objid1): break # No method matches given arguments for ObjectId.op_Equality: (<class 'NoneType'>) == 比较出错
#         objid2 = acad.EntSel(string="请点击第2条对象:")
#         if acad.IsNone(objid2): break 
#         pt1, pd1 = acad.GetStartFinalPoint(objid1)
#         lengthlist1 = acad.GetLWPolyLineLengthList(objid1)
#         pt2, pd2 = acad.GetStartFinalPoint(objid2)
#         lengthlist2 = acad.GetLWPolyLineLengthList(objid2)

#         dy1 = acad.Direct(pt1, pd1)
#         dy1 = acad.Vec3ResetLength(dy1, 2000)
#         pt1 = acad.Vec3Add(pt1, dy1)
#         pt2 = acad.Vec3Add(pt2, dy1)
#         pd1 = acad.Vec3Add(pd1, dy1)
#         pd2 = acad.Vec3Add(pd2, dy1)
#         db1 = acad.Direct(pt1, pd1)
#         db2 = acad.Direct(pt2, pd2)
#         pe1 = acad.Vec3Add(pd1, db1)
#         pe2 = acad.Vec3Add(pd2, db2)
#         linelist = [[pt1, pe1, "0"], [pt2, pe2, "0"]]

#         pt1, pt2 = acad.GetAttachWDirectPt1Pt2(pt1, pt2, 150)
#         linelist.append([pt1, pt2, "0"])

#         for i, [length1, length2] in enumerate(zip(lengthlist1, lengthlist2)):
#             if abs(length1-length2) < 0.0001:
#                 dr1 = acad.GetPerDirectResetLengthXY(pt1, pt2, pe1, length1)
#                 po1 = acad.Vec3Add(pt1, dr1)
#                 po2 = acad.Vec3Add(pt2, dr1)
#             else:
#                 dr1 = acad.GetPerDirectXY(pt1, pt2, pe1)
#                 dd1 = acad.Vec3ResetLength(dr1, length1)
#                 dd2 = acad.Vec3ResetLength(dr1, length2)
#                 po1 = acad.Vec3Add(pt1, dd1)
#                 po2 = acad.Vec3Add(pt2, dd2)

#             linelist.append([pt1, po1, "0"])
#             linelist.append([pt2, po2, "0"])
#             linelist.append([po1, po2, "图层1"])

#             if i == 14:
#                 dr1 = acad.Vec3ResetLength(db1, 200)
#                 pt1 = acad.Vec3Add(po1, dr1)
#                 pt2 = acad.Vec3Add(po2, dr1)
#                 linelist.append([pt1, pt2, "图层1"])
#             else:
#                 pt1, pt2 = po1, po2

#         with acad.transaction() as trans:
#             for pt1, pt2, lname in linelist[:-1]:
#                 acad.AddLine(pt1, pt2, lname)
#             pt1, pt2, lname = linelist[-1]
#             acad.AddLine(pt1, pt2)


# @acad.decorator_command
# def llconnet_rect_and_rect_simple():
#     objid1 = acad.EntSel(string="请点击第1条对象:") 
#     if acad.IsNone(objid1): return
#     objid2 = acad.EntSel(string="请点击第2条对象:")
#     if acad.IsNone(objid2): return
#     ptlist1 = acad.GetLWPolyLinePointList(objid1)[0:4]
#     ptlist2 = acad.GetLWPolyLinePointList(objid2)[0:4]
#     with acad.transaction() as trans:
#         for pt1, pt2 in zip(ptlist1, ptlist2):
#             acad.AddLine(pt1, pt2)



# @acad.decorator_command
# def llconnet_rect_and_rect_point():
#     objid1 = acad.EntSel(string="请点击第1条对象:") # [[0, "LWPOLYLINE"]]
#     if acad.IsNone(objid1): return
#     objid2 = acad.EntSel(string="请点击第2条对象:")
#     if acad.IsNone(objid2): return
#     ptlist1 = acad.GetLWPolyLinePointList(objid1)[0:4]
#     ptlist2 = acad.GetLWPolyLinePointList(objid2)[0:4]

#     buflist = []
#     for pt1 in ptlist1:
#         lengthlist = []
#         for pt2 in ptlist2:
#             length = acad.Distance(pt1, pt2)
#             lengthlist.append([pt2, length])
#         lengthlist.sort(key = lambda item: item[1], reverse=False) # 排序规则，reverse = True 降序， reverse = False 升序（默认）
#         pt2 = lengthlist[0][0]
#         buflist.append([pt1, pt2])

#     with acad.transaction() as trans:
#         for pt1, pt2 in buflist:
#             acad.AddLine(pt1, pt2)

@acad.decorator_command
def llconnet_rect_and_rect_point_for():
    objidlist = acad.SSGetIdList([[0, "LWPOLYLINE"]])
    with acad.transaction() as trans:
        buflist = []
        for objid in objidlist:
            area = acad.TransEntityArea(objid)
            ptlist = acad.TransLWPolyLinePointList(objid)[0:4]
            buflist.append([ptlist, area])
    buflist.sort(key = lambda item: item[1], reverse=True) # 排序规则，reverse = True 降序， reverse = False 升序（默认）

    resultlist = []
    count = len(buflist)
    indexlist = []
    for i in range(count):
        if i in indexlist: continue
        ptwlist = buflist[i][0]
        for k in range(i+1, count):
            ptnlist = buflist[k][0]
            if acad.IsRectInRect(ptnlist, ptwlist):
                resultlist.append([ptnlist, ptwlist])
                indexlist.append(i)
                indexlist.append(k)
                break
    
    buflist = []
    for ptnlist, ptwlist in resultlist:    
        for pt1 in ptnlist:
            lengthlist = []
            for pt2 in ptwlist:
                length = acad.Distance(pt1, pt2)
                lengthlist.append([pt2, length])
            lengthlist.sort(key = lambda item: item[1], reverse=False) # 排序规则，reverse = True 降序， reverse = False 升序（默认）
            pt2 = lengthlist[0][0]
            buflist.append([pt1, pt2])

    with acad.transaction() as trans:
        for pt1, pt2 in buflist:
            acad.AddLine(pt1, pt2)


@acad.decorator_command
def llconnet_rect_and_rect_perx_for():
    objidlist = acad.SSGetIdList([[0, "LWPOLYLINE"]])
    with acad.transaction() as trans:
        buflist = []
        for objid in objidlist:
            area = acad.TransEntityArea(objid)
            ptlist = acad.TransLWPolyLinePointList(objid)[0:4]
            buflist.append([ptlist, area])
    buflist.sort(key = lambda item: item[1], reverse=True) # 排序规则，reverse = True 降序， reverse = False 升序（默认）

    resultlist = []
    count = len(buflist)
    indexlist = []
    for i in range(count):
        if i in indexlist: continue
        ptwlist = buflist[i][0]
        for k in range(i+1, count):
            ptnlist = buflist[k][0]
            if acad.IsRectInRect(ptnlist, ptwlist):
                resultlist.append([ptnlist, ptwlist])
                indexlist.append(i)
                indexlist.append(k)
                break
    
    buflist = []
    for ptnlist, ptwlist in resultlist:    
        for pt1 in ptnlist:
            lengthlist = []
            for pt2 in ptwlist:
                length = acad.Distance(pt1, pt2)
                lengthlist.append([pt2, length])
            lengthlist.sort(key = lambda item: item[1], reverse=False) # 排序规则，reverse = True 降序， reverse = False 升序（默认）
            pt2 = lengthlist[0][0]
            buflist.append([pt1, pt2])


    resultlist = []
    for ptn, ptw in buflist:
        pt1 = ptn
        dx, dy, dz = acad.Direct(ptn, ptw)
        pt2 = acad.Vec3Add(pt1, [dx, 0, 0])
        resultlist.append([pt1, pt2])

    with acad.transaction() as trans:
        for pt1, pt2 in resultlist:
            acad.AddLine(pt1, pt2)





@acad.decorator_command
def llconnet_rect_and_rect_pery_for():
    objidlist = acad.SSGetIdList([[0, "LWPOLYLINE"]])
    with acad.transaction() as trans:
        buflist = []
        for objid in objidlist:
            area = acad.TransEntityArea(objid)
            ptlist = acad.TransLWPolyLinePointList(objid)[0:4]
            buflist.append([ptlist, area])
    buflist.sort(key = lambda item: item[1], reverse=True) # 排序规则，reverse = True 降序， reverse = False 升序（默认）

    resultlist = []
    count = len(buflist)
    indexlist = []
    for i in range(count):
        if i in indexlist: continue
        ptwlist = buflist[i][0]
        for k in range(i+1, count):
            ptnlist = buflist[k][0]
            if acad.IsRectInRect(ptnlist, ptwlist):
                resultlist.append([ptnlist, ptwlist])
                indexlist.append(i)
                indexlist.append(k)
                break
    
    buflist = []
    for ptnlist, ptwlist in resultlist:    
        for pt1 in ptnlist:
            lengthlist = []
            for pt2 in ptwlist:
                length = acad.Distance(pt1, pt2)
                lengthlist.append([pt2, length])
            lengthlist.sort(key = lambda item: item[1], reverse=False) # 排序规则，reverse = True 降序， reverse = False 升序（默认）
            pt2 = lengthlist[0][0]
            buflist.append([pt1, pt2])


    resultlist = []
    for ptn, ptw in buflist:
        pt1 = ptn
        dx, dy, dz = acad.Direct(ptn, ptw)
        pt2 = acad.Vec3Add(pt1, [0, dy, 0])
        resultlist.append([pt1, pt2])

    with acad.transaction() as trans:
        for pt1, pt2 in resultlist:
            acad.AddLine(pt1, pt2)








