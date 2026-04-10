# Convex Hull 凸包算法
import acad

def convex_hull_n3(ptlist):
    count = len(ptlist)
    indexlist = []
    result = []
    for i in range(0, count-1):
        for k in range(i, count):
            pt1 = ptlist[i]
            pt2 = ptlist[k]
            flaglist = []
            for n in range(count):
                if n == i: continue
                if n == k: continue
                pt3 = ptlist[n]
                perflag = acad.GetPerflagXY(pt1, pt2, pt3)
                flaglist.append(perflag)
            if 0 in flaglist: continue
            flagset = set(flaglist)
            if len(flagset) == 1:
                result.append([pt1, pt2])
                indexlist.append(i)
                indexlist.append(k)
    return result


def loopsort(listpoint2):
    count  = len(listpoint2)
    buflist = listpoint2[0:1]
    reflist = listpoint2[1:]
    indexlist = []
    for i in range(count-1):
        pt1, pt2 = buflist[-1]
        for k, [po1, po2] in enumerate(reflist):
            if k in indexlist: continue
            flag = False
            if acad.IsPointSame(pt2, po1):
                flag = True
                buflist.append([po1, po2])
            if acad.IsPointSame(pt2, po2):
                flag = True
                buflist.append([po2, po1])
            if flag:
                indexlist.append(k)
                break
    return buflist


def calc_bound_from_ptlist(ptlist):
    x_min, y_min, z_min = x_max, y_max, z_max = ptlist[0]
    for x1, y1, z1 in ptlist[1:]:
        if x1 < x_min: x_min = x1
        if x1 > x_max: x_max = x1
        if y1 < y_min: y_min = y1
        if y1 > y_max: y_max = y1
        # if z1 < z_min: z_min = z1
        # if z1 > z_max: z_max = z1
    return [x_min, y_min, 0], [x_max, y_min, 0], [x_max, y_max, 0], [x_min, y_max, 0], x_max-x_min, y_max-y_min

def minibound_from_listpoint2(listpoint2):
    ptlist0 = [pt1 for pt1, pt2 in listpoint2]
    result = []
    for pt1, pt2 in listpoint2:
        pt0 = pt1
        dr1 = [1,0,0]
        dr2 = acad.Direct(pt1, pt2)
        angle = acad.AngleFromDotDr1Dr2(dr1, dr2)
        axis = acad.CrossNormalized(dr1, dr2)
        ptlist1 = acad.MatrixRotationPointList(angle, axis, pt0, ptlist0)
        pt1, pt2, pt3, pt4, length, width = calc_bound_from_ptlist(ptlist1)
        po1, po2, po3, po4 = acad.MatrixRotationPointList(-angle, axis, pt0, [pt1, pt2, pt3, pt4])
        result.append([length*width, po1, po2, po3, po4])
    result.sort(key = lambda item: item[0], reverse=False) # 从小到大
    return result[0][1:]