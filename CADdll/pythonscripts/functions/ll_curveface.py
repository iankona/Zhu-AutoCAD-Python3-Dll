import clr

import acad
import academit

import System
import math

def 命令(): 
    # academit.添加命令("llcurve-face-flatten-zhiwen_up", llcurve_face_flatten_zhiwen_up)
    academit.添加命令("llcurve-face-flatten-zhiwen_down", llcurve_face_flatten_zhiwen_down)
    academit.添加命令("llcurve-face-flatten-zhiwen-up-sample", llcurve_face_flatten_zhiwen_up_sample)
    academit.添加命令("llcurve-face-flatten-zhiwen-down-sample", llcurve_face_flatten_zhiwen_down_sample)



llzhu_flatten_sublength = 50
def zhuui_flatten_sublength():
    global llzhu_flatten_sublength
    sublength = acad.GetDouble(llzhu_flatten_sublength, "请输入展平精度(单位mm):")
    if sublength == None: return 
    llzhu_flatten_sublength = sublength





def trans_line_ptlist(pt1, pt2, subcount):
    line1 = acad.DBObjectLine(pt1, pt2)
    sublength = line1.Length/subcount
    ptlist = [pt1]
    sumlength = 0
    for i in range(subcount-1):
        sumlength += sublength
        point = line1.GetPointAtDist(sumlength)
        ptlist.append([point.X, point.Y, point.Z])
    ptlist.append(pt2)
    return ptlist

def trans_line_gap_ptlist(pt1, pt2, subcount):  
    line1 = acad.DBObjectLine(pt1, pt2)
    sublength = line1.Length/subcount
    ptlist = [pt1]
    sumlength = sublength / 2
    for i in range(subcount-1):
        point = line1.GetPointAtDist(sumlength)
        ptlist.append([point.X, point.Y, point.Z])
        sumlength += sublength
    ptlist.append(pt2)
    return ptlist




def trans_line_sublength(pt1, pt2, subcount):
    distance = acad.Distance(pt1, pt2)
    sublength = distance / subcount
    return sublength


def trans_calc_row_column_curve_pointlist(pt1, pt2, pt3, pt4, countrow, countcol):
    ptlist1 = trans_line_ptlist(pt1, pt2, countrow)
    ptlist3 = trans_line_ptlist(pt4, pt3, countrow)

    pointlist = []
    for po1, po2 in zip(ptlist1, ptlist3):
        collist = trans_line_ptlist(po1, po2, countcol)
        pointlist.append(collist)

    looplist = []
    for i in range(countrow):
        for j in range(countcol):
            index1 = [i, j]
            index2 = [i+1, j]
            index3 = [i+1, j+1]
            index4 = [i, j+1]
            if i == 0: 
                looplist.append([index1, index2, index3])
                looplist.append([index1, index3, index4])
            else:
                looplist.append([index1, index3, index4])
                looplist.append([index1, index2, index3])

    return pointlist, looplist

def trans_calc_row_column_curve_gap_pointlist(pt1, pt2, pt3, pt4, countrow, countcol):
    ptlist1 = trans_line_ptlist(pt1, pt2, countrow)
    ptlist3 = trans_line_ptlist(pt4, pt3, countrow)

    pointlist = []
    count = 0
    for po1, po2 in zip(ptlist1, ptlist3):
        count += 1
        if count % 2 == 1:
            collist = trans_line_ptlist(po1, po2, countcol)
        else:
            collist = trans_line_gap_ptlist(po1, po2, countcol)
        pointlist.append(collist)

    looplist = []
    for i in range(countrow):
        for j in range(countcol):
            index1 = [i, j]
            index2 = [i+1, j]
            index3 = [i+1, j+1]
            index4 = [i, j+1]
            if i == 0: 
                looplist.append([index1, index2, index3])
                looplist.append([index1, index3, index4])
            else:
                looplist.append([index1, index3, index4])
                looplist.append([index1, index2, index3])

    return pointlist, looplist




def pointlist_to_lengthlist(pointlist, looplist):
    lengthlist = []
    for loop in looplist: 
        a, b = loop[0]
        c, d = loop[1]
        e, f = loop[2]
        pt1 = pointlist[a][b]
        pt2 = pointlist[c][d]
        pt3 = pointlist[e][f]
        dis1 = acad.Distance(pt1, pt2)
        dis2 = acad.Distance(pt2, pt3)
        dis3 = acad.Distance(pt3, pt1)
        lengthlist.append([dis1, dis2, dis3])
    return lengthlist


def pointlist_to_anglelist(pointlist, looplist):
    anglelist = []
    for loop in looplist: 
        a, b = loop[0]
        c, d = loop[1]
        e, f = loop[2]
        pt1 = pointlist[a][b]
        pt2 = pointlist[c][d]
        pt3 = pointlist[e][f]
        dr1 = acad.Direct(pt1, pt3)
        dr2 = acad.Direct(pt1, pt2)
        dr3 = acad.Direct(pt2, pt1)
        dr4 = acad.Direct(pt2, pt3)
        angle1 = acad.AngleFromDotDr1Dr2(dr1, dr2)
        angle2 = acad.AngleFromDotDr1Dr2(dr3, dr4)
        angle3 = 180-angle1-angle2
        anglelist.append([angle1, angle2, angle3])
    return anglelist



def is_index_in_row_column_index(index, countrow, countcol):
    i, j = index
    if 0 <= i and i < countrow + 1 and 0 <= j and j < countcol + 1: return True
    return False

def row_column_count_to_springloop(countrow, countcol):
    springlooplist = []
    for i in range(countrow+1):
        for j in range(countcol+1):
            index2 = [i+1, j+1]
            index3 = [i,   j+1]
            index4 = [i-1, j+1]

            index1 = [i+1, j]
            index0 = [i  , j]
            index5 = [i-1, j]

            index8 = [i+1, j-1]
            index7 = [i  , j-1]
            index6 = [i-1, j-1]

            loop = []
            if is_index_in_row_column_index(index1, countrow, countcol): loop.append([index0, index1])
            if is_index_in_row_column_index(index2, countrow, countcol): loop.append([index0, index2])
            if is_index_in_row_column_index(index3, countrow, countcol): loop.append([index0, index3])
            if is_index_in_row_column_index(index4, countrow, countcol): loop.append([index0, index4])
            if is_index_in_row_column_index(index5, countrow, countcol): loop.append([index0, index5])
            if is_index_in_row_column_index(index6, countrow, countcol): loop.append([index0, index6])
            if is_index_in_row_column_index(index7, countrow, countcol): loop.append([index0, index7])
            if is_index_in_row_column_index(index8, countrow, countcol): loop.append([index0, index8])
            springlooplist.append(loop)
    return springlooplist

def pointlist_to_springlength(pointlist, springlooplist):
    springlengthlist = []
    for loop in springlooplist:
        lengthlist = []
        for index1, index2 in loop:
            a, b = index1
            c, d = index2
            pt1 = pointlist[a][b]
            pt2 = pointlist[c][d]
            distance = acad.Distance(pt1, pt2)
            lengthlist.append(distance)
        springlengthlist.append(lengthlist)
    return springlengthlist



def spring_update_pointlist(pointlist, springlooplist, springlengthlist):
    result = pointlist[:]
    has_update = False
    for loop, lengthlist in zip(springlooplist, springlengthlist):
        drx = [0,0,0]
        for [index1, index2], length1 in zip(loop, lengthlist):
            a, b = index1
            c, d = index2
            pt1 = pointlist[a][b]
            pt2 = pointlist[c][d]
            distance = acad.Distance(pt1, pt2)
            x = distance - length1
            dx = x/7
            limit = 10
            if abs(dx) > limit:
                if dx > 0: dx = limit
                else: dx = -limit
            dr1 = acad.Direct(pt1, pt2)
            dr1 = acad.Vec3ResetLength(dr1, dx)
            drx = acad.Vec3Add(drx, dr1)
            # print([index1, x, dx, dr1, distance, length1])
        dis = acad.Distance([0,0,0], drx)
        # print(["bias: ", dis])
        if dis < 0.0001: continue
        has_update = True
        pt1 = acad.Vec3Add(pt1, drx)
        result[a][b] = pt1
    return has_update, result


def spring_update_pointlistV2(pointlist, springlooplist, springlengthlist):
    result = pointlist[:]
    has_update = False
    for loop, lengthlist in zip(springlooplist, springlengthlist):
        drx = [0,0,0]
        for [index1, index2], length1 in zip(loop, lengthlist):
            a, b = index1
            c, d = index2
            pt1 = pointlist[a][b]
            pt2 = pointlist[c][d]
            distance = acad.Distance(pt1, pt2)
            x = distance - length1
            # if x < 0.1:
            #     dx = x
            # else:
            #     dx = x/5
            dx = x / 7
            limit = 10
            if abs(dx) > limit:
                if dx > 0: dx = limit
                else: dx = -limit
            dr1 = acad.Direct(pt1, pt2)
            dr1 = acad.Vec3ResetLength(dr1, dx)
            drx = acad.Vec3Add(drx, dr1)
            # print([index1, x, dx, dr1, distance, length1])
        dis = acad.Distance([0,0,0], drx)
        # print(["bias: ", dis])
        # has_update = True
        pt1 = acad.Vec3Add(pt1, drx)
        result[a][b] = pt1
    return has_update, result



def spring_update_pointlist_huke(pointlist, springlooplist, springlengthlist):
    elasticity = 0.7
    result = pointlist[:]
    has_update = False
    for loop, lengthlist in zip(springlooplist, springlengthlist):
        for [index1, index2], length1 in zip(loop, lengthlist):
            a, b = index1
            c, d = index2
            pt1 = result[a][b]
            pt2 = result[c][d]
            dr1 = acad.Direct(pt1, pt2)
            dr2 = acad.Direct(pt2, pt1)
            distance = acad.Distance(pt1, pt2)
            x = distance - length1
            dx = elasticity*x
            dr1 = acad.Vec3ResetLength(dr1, 0.5*dx)
            dr2 = acad.Vec3ResetLength(dr2, 0.5*dx)
            pt1 = acad.Vec3Add(pt1, dr1)
            pt2 = acad.Vec3Add(pt2, dr2)
            result[a][b] = pt1
            result[c][d] = pt2
    return has_update, result
    # //胡克定律：弹力=弹力系数*距离*方向
    # Vector3 vector = positionA - positionB;
    # float distance = vector.magnitude - spring.length; //距离
    # Vector3 direction = vector.normalized; //方向
    # Vector3 move = elasticity * distance * direction;//弹力系数为[0,2]，无阻力时应小于1

    # pointB.position += 0.5f * move;
    # pointA.position += 0.5f * -move;


def lengthlist_deviation(lengthlist1, lengthlist2):
    biaslist = []
    for loop1, loop2 in zip(lengthlist1, lengthlist2):
        dis1, dis2, dis3 = loop1
        len1, len2, len3 = loop2
        bias1 = dis1 - len1 
        bias2 = dis2 - len2 
        bias3 = dis3 - len3
        biaslist.append([bias1, bias2, bias3])


    sumbias = 0
    for bias1, bias2, bias3 in biaslist:
        sumbias += (abs(bias1) + abs(bias2) + abs(bias3))

    # acad.Prompt("长度误差："), acad.Prompt(sumbias)
    # print("长度误差：", sumbias)
    # for bias1, bias2, bias3 in biaslist:
    #     if abs(bias1) < 0.00001: bias1 = 0
    #     if abs(bias2) < 0.00001: bias2 = 0
    #     if abs(bias3) < 0.00001: bias3 = 0
    #     print([bias1, bias2, bias3])
    return sumbias, biaslist



def calc_base_rad(pt1, pt2, pt3, pt4):
    dr1 = acad.Direct(pt1, pt2)
    dr2 = acad.Direct(pt1, pt4)
    angle = acad.AngleFromDotDr1Dr2(dr1, dr2)
    rad = acad.Angle2Rad(angle)
    return rad

def trans_calc_triangle_point(perflag, pt1, dis1, pt2, pt3, dis3):
    circle1 = acad.DBObjectCircle(pt2, dis1)
    circle2 = acad.DBObjectCircle(pt3, dis3)
    ptlist = []
    collect = acad.Point3dCollection() 
    circle1.IntersectWith(circle2, acad.Intersect.OnBothOperands, collect, System.IntPtr.Zero, System.IntPtr.Zero)
    for point in collect: 
        ptlist.append([point.X, point.Y, 0])
    po1, po2 = ptlist
    if acad.GetPerflagXY(pt2, pt3, po1) == perflag: pt1 = po1
    if acad.GetPerflagXY(pt2, pt3, po2) == perflag: pt1 = po2
    return pt1


def trans_calc_triangle_angle_point(perflag, pt1, pt2, pt3, angle2, angle3):
    line0 = acad.DBObjectLine(pt2, pt3)
    if perflag > 0:
        line1 = acad.DBObjectRoationCopy(line0,  angle2, [0,0,1], pt2)
        line2 = acad.DBObjectRoationCopy(line0, -angle3, [0,0,1], pt3)
    else:
        line1 = acad.DBObjectRoationCopy(line0, -angle2, [0,0,1], pt2)
        line2 = acad.DBObjectRoationCopy(line0,  angle3, [0,0,1], pt3)
    ptlist = []
    collect = acad.Point3dCollection() 
    line1.IntersectWith(line2, acad.Intersect.ExtendBoth, collect, System.IntPtr.Zero, System.IntPtr.Zero)
    for point in collect: 
        ptlist.append([point.X, point.Y, 0])
    pt1 = ptlist[0]
    return pt1



def trans_calc_row_column_flatten_pointlist_shape(perflag, pointlist, looplist, lengthlist): # 无约束展开，角长度过长
    pt1 = pointlist[0][0]
    for loop in looplist: 
        a, b = loop[0]
        c, d = loop[1]
        e, f = loop[2]
        pointlist[a][b] = None
        pointlist[c][d] = None
        pointlist[e][f] = None
    dis1 = lengthlist[0][0]
    pt1 = [pt1[0], pt1[1], 0]
    pt2 = acad.Vec3Add(pt1, [dis1, 0, 0])
    a, b = looplist[0][0]
    c, d = looplist[0][1]
    e, f = looplist[0][2]
    pointlist[a][b] = pt1
    pointlist[c][d] = pt2

    for loop, length in zip(looplist, lengthlist): 
        a, b = loop[0]
        c, d = loop[1]
        e, f = loop[2]
        dis1 = length[0]
        dis2 = length[1]
        dis3 = length[2]
        pt1 = pointlist[a][b]
        pt2 = pointlist[c][d]
        pt3 = pointlist[e][f]
        if pt1 == None:
            # print("dis2", acad.Distance(pt2, pt3), "pt1", dis2, pt1, dis1, pt2, pt3, dis3)
            pt1 = trans_calc_triangle_point(perflag, pt1, dis1, pt2, pt3, dis3)
            pointlist[a][b] = pt1
        if pt2 == None:
            # print("dis3", acad.Distance(pt3, pt1), "pt2", dis3, pt2, dis2, pt3, pt1, dis1)
            pt2 = trans_calc_triangle_point(perflag, pt2, dis2, pt3, pt1, dis1)
            pointlist[c][d] = pt2
        if pt3 == None:
            # print("dis1", acad.Distance(pt1, pt2), "pt3", dis1, pt3, dis3, pt1, pt2, dis2)
            pt3 = trans_calc_triangle_point(perflag, pt3, dis3, pt1, pt2, dis2)
            pointlist[e][f] = pt3



def trans_calc_row_column_flatten_pointlist_shape_bias(perflag, pointlist, looplist, lengthlist): # Error, 梯度爆炸
    pt1 = pointlist[0][0]
    for loop in looplist: 
        a, b = loop[0]
        c, d = loop[1]
        e, f = loop[2]
        pointlist[a][b] = None
        pointlist[c][d] = None
        pointlist[e][f] = None
    dis1 = lengthlist[0][0]
    pt1 = [pt1[0], pt1[1], 0]
    pt2 = acad.Vec3Add(pt1, [dis1, 0, 0])
    a, b = looplist[0][0]
    c, d = looplist[0][1]
    e, f = looplist[0][2]
    pointlist[a][b] = pt1
    pointlist[c][d] = pt2

    for loop, length in zip(looplist, lengthlist): 
        a, b = loop[0]
        c, d = loop[1]
        e, f = loop[2]
        dis1 = length[0]
        dis2 = length[1]
        dis3 = length[2]
        pt1 = pointlist[a][b]
        pt2 = pointlist[c][d]
        pt3 = pointlist[e][f]
        if pt1 == None:
            bias = acad.Distance(pt2, pt3)/dis2
            # print("dis2", acad.Distance(pt2, pt3), "pt1", dis2, pt1, dis1, pt2, pt3, dis3)
            pt1 = trans_calc_triangle_point(perflag, pt1, bias*dis1, pt2, pt3, bias*dis3)
            pointlist[a][b] = pt1
        if pt2 == None:
            bias = acad.Distance(pt3, pt1)/dis3
            # print("dis3", acad.Distance(pt3, pt1), "pt2", dis3, pt2, dis2, pt3, pt1, dis1)
            pt2 = trans_calc_triangle_point(perflag, pt2, bias*dis2, pt3, pt1, bias*dis1)
            pointlist[c][d] = pt2
        if pt3 == None:
            bias = acad.Distance(pt1, pt2)/dis1
            # print("dis1", acad.Distance(pt1, pt2), "pt3", dis1, pt3, dis3, pt1, pt2, dis2)
            pt3 = trans_calc_triangle_point(perflag, pt3, bias*dis3, pt1, pt2, bias*dis2)
            pointlist[e][f] = pt3



def trans_calc_row_column_flatten_pointlist_shape_junfen(perflag, pointlist, looplist, lengthlist): # 有限约束展开
    pt1 = pointlist[0][0]
    for loop in looplist: 
        a, b = loop[0]
        c, d = loop[1]
        e, f = loop[2]
        pointlist[a][b] = None
        pointlist[c][d] = None
        pointlist[e][f] = None
    dis1 = lengthlist[0][0]
    pt1 = [pt1[0], pt1[1], 0]
    pt2 = acad.Vec3Add(pt1, [dis1, 0, 0])
    a, b = looplist[0][0]
    c, d = looplist[0][1]
    e, f = looplist[0][2]
    pointlist[a][b] = pt1
    pointlist[c][d] = pt2

    for loop, length in zip(looplist, lengthlist): 
        a, b = loop[0]
        c, d = loop[1]
        e, f = loop[2]
        dis1 = length[0]
        dis2 = length[1]
        dis3 = length[2]
        pt1 = pointlist[a][b]
        pt2 = pointlist[c][d]
        pt3 = pointlist[e][f]
        if pt1 == None:
            dis2 = (acad.Distance(pt2, pt3) + dis2) / 2
            dr2 = acad.Direct(pt2, pt3)
            dr2 = acad.Vec3ResetLength(dr2, dis2)
            pt3 = acad.Vec3Add(pt2, dr2)
            # print("dis2", acad.Distance(pt2, pt3), "pt1", dis2, pt1, dis1, pt2, pt3, dis3)
            pt1 = trans_calc_triangle_point(perflag, pt1, dis1, pt2, pt3, dis3)
            pointlist[a][b] = pt1
            pointlist[e][f] = pt3
        if pt2 == None:
            dis3 = (acad.Distance(pt3, pt1) + dis3) / 2
            dr3 = acad.Direct(pt3, pt1)
            dr3 = acad.Vec3ResetLength(dr3, dis3)
            pt1 = acad.Vec3Add(pt3, dr3)
            # print("dis3", acad.Distance(pt3, pt1), "pt2", dis3, pt2, dis2, pt3, pt1, dis1)
            pt2 = trans_calc_triangle_point(perflag, pt2, dis2, pt3, pt1, dis1)
            pointlist[a][b] = pt1
            pointlist[c][d] = pt2
        if pt3 == None:
            dis1 = (acad.Distance(pt1, pt2) + dis1) / 2
            dr1 = acad.Direct(pt1, pt2)
            dr1 = acad.Vec3ResetLength(dr1, dis1)
            pt2 = acad.Vec3Add(pt1, dr1)
            # print("dis1", acad.Distance(pt1, pt2), "pt3", dis1, pt3, dis3, pt1, pt2, dis2)
            pt3 = trans_calc_triangle_point(perflag, pt3, dis3, pt1, pt2, dis2)
            pointlist[c][d] = pt2
            pointlist[e][f] = pt3





def trans_calc_row_column_flatten_pointlist_shape_limit(perflag, pointlist, looplist, lengthlist): # Error, 梯度爆炸
    pt1 = pointlist[0][0]
    for loop in looplist: 
        a, b = loop[0]
        c, d = loop[1]
        e, f = loop[2]
        pointlist[a][b] = None
        pointlist[c][d] = None
        pointlist[e][f] = None
    dis1 = lengthlist[0][0]
    pt1 = [pt1[0], pt1[1], 0]
    pt2 = acad.Vec3Add(pt1, [dis1, 0, 0])
    a, b = looplist[0][0]
    c, d = looplist[0][1]
    e, f = looplist[0][2]
    pointlist[a][b] = pt1
    pointlist[c][d] = pt2

    for loop, length in zip(looplist, lengthlist): 
        a, b = loop[0]
        c, d = loop[1]
        e, f = loop[2]
        dis1 = length[0]
        dis2 = length[1]
        dis3 = length[2]
        pt1 = pointlist[a][b]
        pt2 = pointlist[c][d]
        pt3 = pointlist[e][f]
        if pt1 == None:
            # dis2 = (acad.Distance(pt2, pt3) + dis2) / 2
            dr2 = acad.Direct(pt2, pt3)
            dr2 = acad.Vec3ResetLength(dr2, dis2)
            pt3 = acad.Vec3Add(pt2, dr2)
            # print("dis2", acad.Distance(pt2, pt3), "pt1", dis2, pt1, dis1, pt2, pt3, dis3)
            pt1 = trans_calc_triangle_point(perflag, pt1, dis1, pt2, pt3, dis3)
            pointlist[a][b] = pt1
            pointlist[e][f] = pt3
        if pt2 == None:
            # dis3 = (acad.Distance(pt3, pt1) + dis3) / 2
            dr3 = acad.Direct(pt3, pt1)
            dr3 = acad.Vec3ResetLength(dr3, dis3)
            pt1 = acad.Vec3Add(pt3, dr3)
            # print("dis3", acad.Distance(pt3, pt1), "pt2", dis3, pt2, dis2, pt3, pt1, dis1)
            pt2 = trans_calc_triangle_point(perflag, pt2, dis2, pt3, pt1, dis1)
            pointlist[a][b] = pt1
            pointlist[c][d] = pt2
        if pt3 == None:
            # dis1 = (acad.Distance(pt1, pt2) + dis1) / 2
            dr1 = acad.Direct(pt1, pt2)
            dr1 = acad.Vec3ResetLength(dr1, dis1)
            pt2 = acad.Vec3Add(pt1, dr1)
            # print("dis1", acad.Distance(pt1, pt2), "pt3", dis1, pt3, dis3, pt1, pt2, dis2)
            pt3 = trans_calc_triangle_point(perflag, pt3, dis3, pt1, pt2, dis2)
            pointlist[c][d] = pt2
            pointlist[e][f] = pt3




def trans_calc_row_column_flatten_pointlist_angle(perflag, pointlist, looplist, lengthlist, anglelist): # Error, 梯度爆炸
    pt1 = pointlist[0][0]
    for loop in looplist: 
        a, b = loop[0]
        c, d = loop[1]
        e, f = loop[2]
        pointlist[a][b] = None
        pointlist[c][d] = None
        pointlist[e][f] = None
    dis1 = lengthlist[0][0]
    pt1 = [pt1[0], pt1[1], 0]
    pt2 = acad.Vec3Add(pt1, [dis1, 0, 0])
    a, b = looplist[0][0]
    c, d = looplist[0][1]
    e, f = looplist[0][2]
    pointlist[a][b] = pt1
    pointlist[c][d] = pt2

    for loop, length in zip(looplist, anglelist): 
        a, b = loop[0]
        c, d = loop[1]
        e, f = loop[2]
        angle1 = length[0]
        angle2 = length[1]
        angle3 = length[2]
        pt1 = pointlist[a][b]
        pt2 = pointlist[c][d]
        pt3 = pointlist[e][f]
        if pt1 == None:
            # print("dis2", acad.Distance(pt2, pt3), "pt1", dis2, pt1, dis1, pt2, pt3, dis3)
            pt1 = trans_calc_triangle_angle_point(perflag, pt1, pt2, pt3, angle2, angle3)
            pointlist[a][b] = pt1
        if pt2 == None:
            # print("dis3", acad.Distance(pt3, pt1), "pt2", dis3, pt2, dis2, pt3, pt1, dis1)
            pt2 = trans_calc_triangle_angle_point(perflag, pt2, pt3, pt1, angle3, angle1)
            pointlist[c][d] = pt2
        if pt3 == None:
            # print("dis1", acad.Distance(pt1, pt2), "pt3", dis1, pt3, dis3, pt1, pt2, dis2)
            pt3 = trans_calc_triangle_angle_point(perflag, pt3, pt1, pt2, angle1, angle2)
            pointlist[e][f] = pt3




@acad.decorator_command
def llcurve_face_flatten_zhiwen_down():
    zhuui_flatten_sublength()
    pt1, pt2, pt3, pt4 = acad.GetPoint4()
    # print("pt1, pt2, pt3, pt4 =", [pt1, pt2, pt3, pt4])
    if pt1 == None: return
    dis1 = acad.Distance(pt1, pt2)
    dis2 = acad.Distance(pt2, pt3)
    dis3 = acad.Distance(pt3, pt4)
    dis4 = acad.Distance(pt4, pt1)
    maxrowlength = max(dis1, dis3)
    maxcollength = max(dis2, dis4)
    countrow = int(maxrowlength/llzhu_flatten_sublength)
    countcol = int(maxcollength/llzhu_flatten_sublength)

    pointlist, looplist = trans_calc_row_column_curve_pointlist(pt1, pt2, pt3, pt4, countrow, countcol)
    lengthlist = pointlist_to_lengthlist(pointlist, looplist)
    # anglelist = pointlist_to_anglelist(pointlist, looplist)
    draw_curve_face_mesh(pointlist, looplist)
    perflag = -1

    springlooplist = row_column_count_to_springloop(countrow, countcol)
    springlengthlist = pointlist_to_springlength(pointlist, springlooplist)

    trans_calc_row_column_flatten_pointlist_shape_junfen(perflag, pointlist, looplist, lengthlist)
    draw_curve_face_mesh(pointlist, looplist)
    lengthlist2 = pointlist_to_lengthlist(pointlist, looplist)
    sumbias, biaslist = lengthlist_deviation(lengthlist, lengthlist2)
    acad.Prompt(sumbias), acad.Prompt("\n")

    # print("===========")
    # print(looplist[-2:])
    # print(springlooplist[-1])
    # print("================")


    # print(pointlist, springlooplist, springlengthlist)
    springlooplist_reverse = springlooplist[::-1]
    springlengthlist_reverse = springlengthlist[::-1]
    bias = sumbias
    count = 0
    while True:
        count += 1
        # print(f"第{count}迭代:")
        flag, pointlist = spring_update_pointlist_huke(pointlist, springlooplist, springlengthlist)
        # if count % 2 == 1: # 效果极差
        #     flag, pointlist = spring_update_pointlist_huke(pointlist, springlooplist, springlengthlist)
        # else:
        #     flag, pointlist = spring_update_pointlist_huke(pointlist, springlooplist_reverse, springlengthlist_reverse)
        lengthlist2 = pointlist_to_lengthlist(pointlist, looplist)    
        sumbias, biaslist = lengthlist_deviation(lengthlist, lengthlist2)  
        if sumbias < bias: bias = sumbias
        # if count % 2 == 1:
        #     flag, pointlist = spring_update_pointlist(pointlist, springlooplist, springlengthlist)
        # else:
        #     flag, pointlist = spring_update_pointlist(pointlist, springlooplist_reverse, springlengthlist_reverse)
        # if flag == False: break



        # print("sumbias:", sumbias)
        if count > 10 and sumbias > bias: break
        if sumbias < 5: break
        if count > 1000: break

    acad.Prompt(sumbias)
      
    # draw_curve_face_nias(pointlist, looplist, biaslist)
    draw_curve_face_mesh(pointlist, looplist)

    # trans_calc_row_column_flatten_pointlist_shape_limit(perflag, pointlist, looplist, lengthlist)
    # draw_curve_face_mesh(pointlist, looplist)
    # lengthlist2 = pointlist_to_lengthlist(pointlist, looplist)
    # lengthlist_deviation(lengthlist, lengthlist2)


    # trans_calc_row_column_flatten_pointlist_angle(perflag, pointlist, looplist, lengthlist, anglelist)
    # draw_curve_face_mesh(pointlist, looplist)
    # lengthlist3 = pointlist_to_lengthlist(pointlist, looplist)
    # lengthlist_deviation(lengthlist, lengthlist3)


def __segment_center(pt1, pt2):
    x1, y1, z1 = pt1
    x2, y2, z2 = pt2
    x0 = (x1+x2)/2
    y0 = (y1+y2)/2
    z0 = (z1+z2)/2
    return [x0, y0, z0]

def __triangle_center(pt1, pt2, pt3):
    x1, y1, z1 = pt1
    x2, y2, z2 = pt2
    x3, y3, z3 = pt3
    x0 = (x1+x2+x3)/3
    y0 = (y1+y2+y3)/3
    z0 = (z1+z2+z3)/3
    return [x0, y0, z0]


def draw_curve_face_mesh(pointlist, looplist): 
    with acad.transaction() as trans:
        objidlist = []
        count = 0
        for loop in looplist: 
            count += 1
            a, b = loop[0]
            c, d = loop[1]
            e, f = loop[2]
            pt1 = pointlist[a][b]
            pt2 = pointlist[c][d]
            pt3 = pointlist[e][f]
            index = count%255
            line1 = acad.AddLine(pt1, pt2, color_index=index)
            line2 = acad.AddLine(pt2, pt3, color_index=index)
            line3 = acad.AddLine(pt3, pt1, color_index=index)
            objidlist.append(line1.ObjectId)
            objidlist.append(line2.ObjectId)
            objidlist.append(line3.ObjectId)
            pt4 = __triangle_center(pt1, pt2, pt3)
            text1 = acad.AddText(pt4, string=str(count), size=5)
            objidlist.append(text1.ObjectId)
        acad.AddGroup(objidlist)



def draw_curve_face_bias(pointlist, looplist, biaslist): 
    with acad.transaction() as trans:
        objidlist = []
        count = 0
        for loop, bias in zip(looplist, biaslist): 
            count += 1
            a, b = loop[0]
            c, d = loop[1]
            e, f = loop[2]
            pt1 = pointlist[a][b]
            pt2 = pointlist[c][d]
            pt3 = pointlist[e][f]
            mid1 = __segment_center(pt1, pt2)
            mid2 = __segment_center(pt2, pt3)
            mid3 = __segment_center(pt3, pt1)
            text2 = acad.AddText(mid1, string=str(bias[0]), size=5)
            text3 = acad.AddText(mid2, string=str(bias[1]), size=5)
            text4 = acad.AddText(mid3, string=str(bias[2]), size=5)
            objidlist.append(text2.ObjectId)
            objidlist.append(text3.ObjectId)
            objidlist.append(text4.ObjectId)

        acad.AddGroup(objidlist)



@acad.decorator_command
def llcurve_face_flatten_zhiwen_up_sample():
    pt1, pt2, pt3, pt4 = acad.GetPoint4()
    if pt1 == None: return
    dis1 = acad.Distance(pt1, pt2)
    dis2 = acad.Distance(pt2, pt3)
    dis3 = acad.Distance(pt3, pt4)
    dis4 = acad.Distance(pt4, pt1)
    dr1 = acad.Direct(pt1, pt2)
    dr2 = acad.Direct(pt2, pt3)
    angle = acad.AngleFromDotDr1Dr2(dr1, dr2)
    rad = acad.Angle2Rad(angle)
    pt1 = pt1[0:2]
    pt2 = acad.Vec3Add(pt1, [dis1,0,0])
    pt3 = acad.Vec3Add(pt2, [dis2*math.cos(rad),  dis2*math.sin(rad), 0]) 
    with acad.transaction() as trans:
        circle1 = acad.DBObjectCircle(pt1, dis4)
        circle2 = acad.DBObjectCircle(pt3, dis3)
        ptlist = []
        collect = acad.Point3dCollection() 
        circle1.IntersectWith(circle2, acad.Intersect.OnBothOperands, collect, System.IntPtr.Zero, System.IntPtr.Zero)
        for point in collect: 
            ptlist.append([point.X, point.Y, 0])
        po1, po2 = ptlist
        perflag = acad.GetPerflagXY(pt1, pt2, pt3)
        if acad.GetPerflagXY(pt1, pt2, po1) == perflag: pt4 = po1
        if acad.GetPerflagXY(pt1, pt2, po2) == perflag: pt4 = po2
        acad.AddMKPolyLine([pt1, pt2, pt3, pt4])



@acad.decorator_command
def llcurve_face_flatten_zhiwen_down_sample():
    pt1, pt2, pt3, pt4 = acad.GetPoint4()
    if pt1 == None: return
    dis1 = acad.Distance(pt1, pt2)
    dis2 = acad.Distance(pt2, pt3)
    dis3 = acad.Distance(pt3, pt4)
    dis4 = acad.Distance(pt4, pt1)
    dr1 = acad.Direct(pt1, pt2)
    dr2 = acad.Direct(pt2, pt3)
    angle = acad.AngleFromDotDr1Dr2(dr1, dr2)
    rad = acad.Angle2Rad(angle)
    pt1 = pt1[0:2]
    pt2 = acad.Vec3Add(pt1, [dis1,0,0])
    pt3 = acad.Vec3Add(pt2, [dis2*math.cos(rad), -dis2*math.sin(rad), 0]) 
    with acad.transaction() as trans:
        circle1 = acad.DBObjectCircle(pt1, dis4)
        circle2 = acad.DBObjectCircle(pt3, dis3)
        ptlist = []
        collect = acad.Point3dCollection() 
        circle1.IntersectWith(circle2, acad.Intersect.OnBothOperands, collect, System.IntPtr.Zero, System.IntPtr.Zero)
        for point in collect: 
            ptlist.append([point.X, point.Y, 0])
        po1, po2 = ptlist
        perflag = acad.GetPerflagXY(pt1, pt2, pt3)
        if acad.GetPerflagXY(pt1, pt2, po1) == perflag: pt4 = po1
        if acad.GetPerflagXY(pt1, pt2, po2) == perflag: pt4 = po2
        acad.AddMKPolyLine([pt1, pt2, pt3, pt4])




def trans_calc_row_column_curve_pointlistV0(pt1, pt2, pt3, pt4, countrow, countcol):
    ptlist1 = trans_line_ptlist(pt1, pt2, countrow)
    ptlist3 = trans_line_ptlist(pt4, pt3, countrow)

    pointlist = []
    for po1, po2 in zip(ptlist1, ptlist3):
        collist = trans_line_ptlist(po1, po2, countcol)
        pointlist.append(collist)

    looplist = []
    for i in range(countrow):
        for j in range(countcol):
            index1 = [i, j]
            index2 = [i+1, j]
            index3 = [i+1, j+1]
            index4 = [i, j+1]
            if i == 0: 
                looplist.append([index1, index2, index3])
                looplist.append([index1, index3, index4])
            else:
                looplist.append([index1, index3, index4])
                looplist.append([index1, index2, index3])


    lengthlist = []
    for loop in looplist: 
        a, b = loop[0]
        c, d = loop[1]
        e, f = loop[2]
        pt1 = pointlist[a][b]
        pt2 = pointlist[c][d]
        pt3 = pointlist[e][f]
        dis1 = acad.Distance(pt1, pt2)
        dis2 = acad.Distance(pt2, pt3)
        dis3 = acad.Distance(pt3, pt1)
        lengthlist.append([dis1, dis2, dis3])


    return pointlist, looplist, lengthlist

