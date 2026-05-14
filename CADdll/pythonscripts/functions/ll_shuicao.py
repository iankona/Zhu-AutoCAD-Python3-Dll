import clr

import acad
import academit

import System

def 命令(): 
    academit.添加命令("llsc-same-for", llsc_same_for)
    academit.添加命令("llsc-diff-for", llsc_diff_for)
    academit.添加命令("llsc-diff-dajie-for", llsc_diff_dajie_for)


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

def __input_string_numblist(char):
    char.replace("　", "")
    char.replace(" ", "")
    char.replace("，", ",")
    if char.startswith(","): char = char[1:]
    if char.endswith(","): char = char[0:-1]
    charlist = char.split(",")
    numblist = [float(char) for char in charlist]
    acad.Prompt(numblist), acad.Prompt("\n")
    return numblist


def __input_string_sumblist(char, length):
    numblist = __input_string_numblist(char)
    sumblist = []
    sum = 0
    for num in numblist:
        sum += num
        if sum < length: sumblist.append(sum)
    return sumblist




def __edgelist_to_flaglist(edgelist):
    count = len(edgelist)
    flaglist = [0]
    for i in range(1, count-1):
        pt1, pt2 = edgelist[i-1]
        pt3, pt4 = edgelist[i]
        pt5, pt6 = edgelist[i+1]
        perflag1 = acad.GetPerflagXY(pt3, pt4, pt1)
        perflag2 = acad.GetPerflagXY(pt3, pt4, pt6)
        flag = 0
        if perflag1 > 0 and perflag2 > 0 : flag = 1
        if perflag1 < 0 and perflag2 < 0 : flag = 1
        flaglist.append(flag)
    flaglist.append(0)
    return flaglist

def __edgelist_to_lenglist(edgelist):
    lenglist = []
    for pt1, pt2 in edgelist:
        distance = acad.Distance(pt1, pt2)
        lenglist.append(distance)
    return lenglist


def __edgelist_to_indexlist(edgelist, mid1):
    indexlist = []
    for i, [pt1, pt2]  in enumerate(edgelist):
        mid2 = acad.MidPt1Pt2(pt1, pt2)
        if acad.IsPointSame(mid1, mid2):
            indexlist = [i-1, i, i+1]
            break
    return indexlist



def __sublengthlist_from_samelist(length, sumblist, lenmin, lenmax):
    sublengthlist = [lenmin]
    for sumb in sumblist:
        sublengthlist.append(lenmin)
    sublengthlist.append(lenmin)    
    return sublengthlist

def __sublengthlist_from_sumblist(length, sumblist, lenmin, lenmax):
    sublengthlist = [lenmin]
    for sumb in sumblist:
        a = sumb / length
        b = 1 - a
        sublength = lenmin*b + lenmax*a
        sublengthlist.append(sublength)
    sublengthlist.append(lenmax)    
    return sublengthlist


def __sublengthlist_subtract_rationess(rationess, sublengthlist, lenmin, lenmax):
    sublist2 = [lenmin]
    for number in sublengthlist[1:-1]:
        sublist2.append(number-rationess)
    sublist2.append(lenmax)
    return sublist2

def __calc_lwpl(length, rationess, sumblist, pd1, edgelist1, edgelist2): 
    flaglist1 = __edgelist_to_flaglist(edgelist1)
    lenglist1 = __edgelist_to_lenglist(edgelist1)
    lenglist2 = __edgelist_to_lenglist(edgelist2)
    indexlist = __edgelist_to_indexlist(edgelist1, pd1)
    buflist = []
    for i, [flag, lenmin, lenmax] in enumerate(zip(flaglist1, lenglist1, lenglist2)):
        if abs(lenmax-lenmin) < 1: 
            sublist1 = __sublengthlist_from_samelist(length, sumblist, lenmin, lenmax)
        else: 
            sublist1 = __sublengthlist_from_sumblist(length, sumblist, lenmin, lenmax)

        if flag and i in indexlist:
            sublist2 = __sublengthlist_subtract_rationess(rationess, sublist1, lenmin, lenmax)
        else:
            sublist2 = sublist1
        buflist.append([sublist1, sublist2])
    result = []
    for j in range(len(sumblist)+1):
        lengthlist1 = []
        lengthlist2 = []
        for sublist1, sublist2 in buflist:
            lengthlist1.append(sublist1[j])
            lengthlist2.append(sublist2[j+1])
        result.append([lengthlist1, lengthlist2])

    maxsumlength = 300
    for length in lenglist2:
        maxsumlength += length
    
    return indexlist[1], maxsumlength, result


def __align_lwpl_from_mid_index(baselist, lengthlist1, lengthlist2, index):
    sumb0 = 0
    sumb1 = 0
    sumb2 = 0
    for i, [length0, length1, length2] in enumerate(zip(baselist, lengthlist1, lengthlist2)):
        if i < index:
            sumb0 += length0
            sumb1 += length1
            sumb2 += length2
        if i == index:
            sumb0 += length0/2
            sumb1 += length1/2
            sumb2 += length2/2
            break
    dy1 = sumb1 - sumb0
    dy2 = sumb2 - sumb0
    return dy1, dy2

def __build_lwpl_ydirect(pt1, lengthlist1):
    ptlist = [pt1]
    for length1 in lengthlist1:
        pt1 = acad.Vec3Add(pt1, [0, -length1, 0])
        ptlist.append(pt1)
    return ptlist


@acad.decorator_command_undo
def llsc_diff_dajie_for():
    length = acad.GetDouble(13000, "请输入水槽总长度:")
    ness = acad.GetDouble(1, "请输入水槽板厚:")
    rationess = (2*ness)+1
    char = acad.GetString("3010,3000,3000,3990", "请输入分段长度字符串")
    numblist = __input_string_numblist(char)
    sumblist = __input_string_sumblist(char, length)
    char = acad.GetString("30,40,40,0", "请输入搭接长度字符串")
    djielist = __input_string_numblist(char)
    pt1, objid1 = acad.EntSelEntity("请选择第1条短多段线:")
    po1, objid2 = acad.EntSelEntity("请选择第2条长多段线:")
    edgelist1 = acad.GetLWPolyLineEdgeList(objid1, pt1)
    edgelist2 = acad.GetLWPolyLineEdgeList(objid2, po1)
    pd1 = acad.GetPoint("请选择对齐点1:")
    if pd1 == None: return 
    index, offsety, buflist = __calc_lwpl(length, rationess, sumblist, pd1, edgelist1, edgelist2)
    pa1 = acad.GetPoint("请选择基点")
    if pa1 == None: return 
    pb1 = acad.Vec3Add(pa1, [0, -offsety])
    

    with acad.transaction() as trans:
        baselist = buflist[0][0]
        for i, [lengthlist1, lengthlist2] in enumerate(buflist):
            length = numblist[i]
            dx = djielist[i]
            dy1, dy2 = __align_lwpl_from_mid_index(baselist, lengthlist1, lengthlist2, index)
            pa2 = acad.Vec3Add(pa1, [length, 0])
            pb2 = acad.Vec3Add(pb1, [length, 0])
            pa1 = acad.Vec3Add(pa1, [0, dy1])
            pb1 = acad.Vec3Add(pb1, [0, dy1])
            pc2 = acad.Vec3Add(pa2, [dx, dy2])
            pd2 = acad.Vec3Add(pb2, [dx, dy2])
            if i % 2 == 0:
                ptlist1 = __build_lwpl_ydirect(pa1, lengthlist1)
                ptlist2 = __build_lwpl_ydirect(pc2, lengthlist2)
            if i % 2 == 1:
                ptlist1 = __build_lwpl_ydirect(pb1, lengthlist1)
                ptlist2 = __build_lwpl_ydirect(pd2, lengthlist2)
            count = len(lengthlist1)
            for i, [pt1, pt2] in enumerate(zip(ptlist1, ptlist2)):
                if i == 0 or  i == count: acad.AddLine(pt1, pt2, "0")
                if i >  0 and i <  count: acad.AddLine(pt1, pt2, "图层1")  
            acad.AddLWPolyLine(ptlist1, "0")     
            acad.AddLWPolyLine(ptlist2, "0")     
            pa1 = pa2
            pb1 = pb2







    



