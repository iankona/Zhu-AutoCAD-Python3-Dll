
from Autodesk.AutoCAD.DatabaseServices import DBPoint, Extents3d, Polyline, Polyline3d, Line, Circle, Poly3dType, DBText, Region, DBObjectCollection, Intersect



import acad
import academit
import System

def 命令(): 
    academit.添加命令("llalign-pmin-x", llalign_pmin_x)
    academit.添加命令("llalign-pmin-y", llalign_pmin_y)
    academit.添加命令("llalign-pmax-x", llalign_pmax_x)
    academit.添加命令("llalign-pmax-y", llalign_pmax_y)
    academit.添加命令("llalign-center-x", llalign_center_x)
    academit.添加命令("llalign-center-y", llalign_center_y)
    academit.添加命令("llalign-pmin-x-for", llalign_pmin_x_for)
    academit.添加命令("llalign-pmin-y-for", llalign_pmin_y_for)
    academit.添加命令("llalign-pmax-x-for", llalign_pmax_x_for)
    academit.添加命令("llalign-pmax-y-for", llalign_pmax_y_for)
    academit.添加命令("llalign-center-x-for", llalign_center_x_for)
    academit.添加命令("llalign-center-y-for", llalign_center_y_for)
    academit.添加命令("llalign-distance-x-for", llalign_distance_x_for)
    academit.添加命令("llalign-distance-y-for", llalign_distance_y_for)

zhu_llalign_distance_x = 500
def zhu_ui_llalign_distance_x():
    global zhu_llalign_distance_x
    distance = acad.GetDouble(zhu_llalign_distance_x, string="请输入X间隔: ")
    if distance != None: zhu_llalign_distance_x = distance

zhu_llalign_distance_y = 500
def zhu_ui_llalign_distance_y():
    global zhu_llalign_distance_y
    distance = acad.GetDouble(zhu_llalign_distance_y, string="请输入Y间隔: ")
    if distance != None: zhu_llalign_distance_y = distance

zhu_llalign_distance = 500
def zhu_ui_llalign_distance():
    global zhu_llalign_distance
    distance = acad.GetDouble(zhu_llalign_distance, string="请输入间隔: ")
    if distance != None: zhu_llalign_distance = distance


@acad.decorator_command
def llalign_pmin_x():
    objidlist = acad.SSGetIdList(string="请选择对齐对象: ")  
    if objidlist == None: return
    # pb1 = acad.GetPoint("请选择对齐基点: ")
    # if pb1 == None: return
    # xb, yb, zb = pb1
    with acad.transaction() as trans:
        pb1, pb2 = acad.TransObjectIdListBoundXY0(objidlist)
        xb, yb, zb = pb1
        result = acad.TransAutoFindRegionRectList(objidlist)
        for pt1, pt2 in result: 
            x1, y1, z1 = pt1
            po1 = [0, y1, 0]
            po2 = [0, yb, 0]
            objidlist = acad.GetSelectCornerCrossIdList(pt1, pt2)
            for objid in objidlist: acad.TransMove(objid, po1, po2)

            
@acad.decorator_command
def llalign_pmin_y():
    objidlist = acad.SSGetIdList(string="请选择对齐对象: ")  
    if objidlist == None: return
    # pb1 = acad.GetPoint("请选择对齐基点: ")
    # if pb1 == None: return
    # xb, yb, zb = pb1
    with acad.transaction() as trans:
        pb1, pb2 = acad.TransObjectIdListBoundXY0(objidlist)
        xb, yb, zb = pb1
        result = acad.TransAutoFindRegionRectList(objidlist)
        for pt1, pt2 in result: 
            x1, y1, z1 = pt1
            po1 = [x1, 0, 0]
            po2 = [xb, 0, 0]
            objidlist = acad.GetSelectCornerCrossIdList(pt1, pt2)
            for objid in objidlist: acad.TransMove(objid, po1, po2)


@acad.decorator_command
def llalign_pmax_x():
    objidlist = acad.SSGetIdList(string="请选择对齐对象: ")  
    if objidlist == None: return
    # pb1 = acad.GetPoint("请选择对齐基点: ")
    # if pb1 == None: return
    # xb, yb, zb = pb1
    with acad.transaction() as trans:
        pb1, pb2 = acad.TransObjectIdListBoundXY0(objidlist)
        xb, yb, zb = pb2
        result = acad.TransAutoFindRegionRectList(objidlist)
        for pt1, pt2 in result: 
            x1, y1, z1 = pt2
            po1 = [0, y1, 0]
            po2 = [0, yb, 0]
            objidlist = acad.GetSelectCornerCrossIdList(pt1, pt2)
            for objid in objidlist: acad.TransMove(objid, po1, po2)
        

@acad.decorator_command
def llalign_pmax_y():
    objidlist = acad.SSGetIdList(string="请选择对齐对象: ")  
    if objidlist == None: return
    # pb1 = acad.GetPoint("请选择对齐基点: ")
    # if pb1 == None: return
    # xb, yb, zb = pb1
    with acad.transaction() as trans:
        pb1, pb2 = acad.TransObjectIdListBoundXY0(objidlist)
        xb, yb, zb = pb2
        result = acad.TransAutoFindRegionRectList(objidlist)
        for pt1, pt2 in result: 
            x1, y1, z1 = pt2
            po1 = [x1, 0, 0]
            po2 = [xb, 0, 0]
            objidlist = acad.GetSelectCornerCrossIdList(pt1, pt2)
            for objid in objidlist: acad.TransMove(objid, po1, po2)


@acad.decorator_command
def llalign_center_x():
    objidlist = acad.SSGetIdList(string="请选择对齐对象: ")  
    if objidlist == None: return
    # pb1 = acad.GetPoint("请选择对齐基点: ")
    # if pb1 == None: return
    # xb, yb, zb = pb1
    with acad.transaction() as trans:
        center = acad.TransObjectIdListBoundCenterXY0(objidlist)
        xb, yb, zb = center
        result = acad.TransAutoFindRegionRectList(objidlist)
        for pt1, pt2 in result: 
            x1, y1, z1 = [(pt1[0]+pt2[0])/2, (pt1[1]+pt2[1])/2, 0]
            po1 = [0, y1, 0]
            po2 = [0, yb, 0]
            objidlist = acad.GetSelectCornerCrossIdList(pt1, pt2)
            for objid in objidlist: acad.TransMove(objid, po1, po2)
     

@acad.decorator_command
def llalign_center_y():
    objidlist = acad.SSGetIdList(string="请选择对齐对象: ")  
    if objidlist == None: return
    # pb1 = acad.GetPoint("请选择对齐基点: ")
    # if pb1 == None: return
    # xb, yb, zb = pb1
    with acad.transaction() as trans:
        center = acad.TransObjectIdListBoundCenterXY0(objidlist)
        xb, yb, zb = center
        result = acad.TransAutoFindRegionRectList(objidlist)
        for pt1, pt2 in result: 
            x1, y1, z1 = [(pt1[0]+pt2[0])/2, (pt1[1]+pt2[1])/2, 0]
            po1 = [x1, 0, 0]
            po2 = [xb, 0, 0]
            objidlist = acad.GetSelectCornerCrossIdList(pt1, pt2)
            for objid in objidlist: acad.TransMove(objid, po1, po2)


@acad.decorator_command
def llalign_pmin_x_for():
    while True:
        objidlist = acad.SSGetIdList(string="请选择对齐对象: ")  
        if objidlist == None: break
        # pb1 = acad.GetPoint("请选择对齐基点: ")
        # if pb1 == None: break
        # xb, yb, zb = pb1
        with acad.transaction() as trans:
            pb1, pb2 = acad.TransObjectIdListBoundXY0(objidlist)
            xb, yb, zb = pb1
            result = acad.TransAutoFindRegionRectList(objidlist)
            for pt1, pt2 in result: 
                x1, y1, z1 = pt1
                po1 = [0, y1, 0]
                po2 = [0, yb, 0]
                objidlist = acad.GetSelectCornerCrossIdList(pt1, pt2)
                for objid in objidlist: acad.TransMove(objid, po1, po2)
       

@acad.decorator_command
def llalign_pmin_y_for():
    while True:
        objidlist = acad.SSGetIdList(string="请选择对齐对象: ")  
        if objidlist == None: break
        # pb1 = acad.GetPoint("请选择对齐基点: ")
        # if pb1 == None: break
        # xb, yb, zb = pb1
        with acad.transaction() as trans:
            pb1, pb2 = acad.TransObjectIdListBoundXY0(objidlist)
            xb, yb, zb = pb1
            result = acad.TransAutoFindRegionRectList(objidlist)
            for pt1, pt2 in result: 
                x1, y1, z1 = pt1
                po1 = [x1, 0, 0]
                po2 = [xb, 0, 0]
                objidlist = acad.GetSelectCornerCrossIdList(pt1, pt2)
                for objid in objidlist: acad.TransMove(objid, po1, po2)


@acad.decorator_command
def llalign_pmax_x_for():
    while True:
        objidlist = acad.SSGetIdList(string="请选择对齐对象: ")  
        if objidlist == None: break
        # pb1 = acad.GetPoint("请选择对齐基点: ")
        # if pb1 == None: break
        # xb, yb, zb = pb1
        with acad.transaction() as trans:
            pb1, pb2 = acad.TransObjectIdListBoundXY0(objidlist)
            xb, yb, zb = pb2
            result = acad.TransAutoFindRegionRectList(objidlist)
            for pt1, pt2 in result: 
                x1, y1, z1 = pt2
                po1 = [0, y1, 0]
                po2 = [0, yb, 0]
                objidlist = acad.GetSelectCornerCrossIdList(pt1, pt2)
                for objid in objidlist: acad.TransMove(objid, po1, po2)


@acad.decorator_command
def llalign_pmax_y_for():
    while True:
        objidlist = acad.SSGetIdList(string="请选择对齐对象: ") 
        if objidlist == None: break 
        # pb1 = acad.GetPoint("请选择对齐基点: ")
        # if pb1 == None: break
        # xb, yb, zb = pb1
        with acad.transaction() as trans:
            pb1, pb2 = acad.TransObjectIdListBoundXY0(objidlist)
            xb, yb, zb = pb2
            result = acad.TransAutoFindRegionRectList(objidlist)
            for pt1, pt2 in result: 
                x1, y1, z1 = pt2
                po1 = [x1, 0, 0]
                po2 = [xb, 0, 0]
                objidlist = acad.GetSelectCornerCrossIdList(pt1, pt2)
                for objid in objidlist: acad.TransMove(objid, po1, po2)


@acad.decorator_command
def llalign_center_x_for():
    while True:
        objidlist = acad.SSGetIdList(string="请选择对齐对象: ") 
        if objidlist == None: break 
        # pb1 = acad.GetPoint("请选择对齐基点: ")
        # if pb1 == None: break
        # xb, yb, zb = pb1
        with acad.transaction() as trans:
            center = acad.TransObjectIdListBoundCenterXY0(objidlist)
            xb, yb, zb = center
            result = acad.TransAutoFindRegionRectList(objidlist)
            for pt1, pt2 in result: 
                x1, y1, z1 = [(pt1[0]+pt2[0])/2, (pt1[1]+pt2[1])/2, 0]
                po1 = [0, y1, 0]
                po2 = [0, yb, 0]
                objidlist = acad.GetSelectCornerCrossIdList(pt1, pt2)
                for objid in objidlist: acad.TransMove(objid, po1, po2)


@acad.decorator_command
def llalign_center_y_for():
    while True:
        objidlist = acad.SSGetIdList(string="请选择对齐对象: ")  
        if objidlist == None: break
        # pb1 = acad.GetPoint("请选择对齐基点: ")
        # if pb1 == None: break
        # xb, yb, zb = pb1
        with acad.transaction() as trans:
            center = acad.TransObjectIdListBoundCenterXY0(objidlist)
            xb, yb, zb = center
            result = acad.TransAutoFindRegionRectList(objidlist)
            for pt1, pt2 in result: 
                x1, y1, z1 = [(pt1[0]+pt2[0])/2, (pt1[1]+pt2[1])/2, 0]
                po1 = [x1, 0, 0]
                po2 = [xb, 0, 0]
                objidlist = acad.GetSelectCornerCrossIdList(pt1, pt2)
                for objid in objidlist: acad.TransMove(objid, po1, po2)



@acad.decorator_command
def llalign_distance_x_for():
    zhu_ui_llalign_distance_x()
    objidlist = acad.SSGetIdList(string="请选择对齐对象: ")  
    if objidlist == None: return
    with acad.transaction() as trans:
        result = acad.TransAutoFindRegionRectList(objidlist)
        buflist = []
        for pt1, pt2 in result:
            x1, y1, z1 = pt1
            x2, y2, z2 = pt2
            length, width = x2-x1, y2-y1
            buflist.append([x1, length, width, [pt1, pt2]])
        buflist.sort(key = lambda item: item[0], reverse=False) # 排序规则，reverse = True 降序， reverse = False 升序（默认）    
        x0 = buflist[0][0]
        for i in range(len(buflist)):
            length = buflist[i][1]
            buflist[i][0] = x0
            x0 = x0 + length + zhu_llalign_distance_x
        for xt, length, width, [pt1, pt2] in buflist:
            x1, y1, z1 = pt1
            po1 = [x1, 0, 0]
            po2 = [xt, 0, 0]
            objidlist = acad.GetSelectCornerCrossIdList(pt1, pt2)
            for objid in objidlist: acad.TransMove(objid, po1, po2)

@acad.decorator_command
def llalign_distance_y_for():
    zhu_ui_llalign_distance_y()
    objidlist = acad.SSGetIdList(string="请选择对齐对象: ")  
    if objidlist == None: return
    with acad.transaction() as trans:
        result = acad.TransAutoFindRegionRectList(objidlist)
        buflist = []
        for pt1, pt2 in result:
            x1, y1, z1 = pt1
            x2, y2, z2 = pt2
            length, width = x2-x1, y2-y1
            buflist.append([y1, length, width, [pt1, pt2]])
        buflist.sort(key = lambda item: item[0], reverse=False) # 排序规则，reverse = True 降序， reverse = False 升序（默认）    
        y0 = buflist[0][0]
        for i in range(len(buflist)):
            width = buflist[i][2]
            buflist[i][0] = y0
            y0 = y0 + width + zhu_llalign_distance_y
        for yt, length, width, [pt1, pt2] in buflist:
            x1, y1, z1 = pt1
            po1 = [0, y1, 0]
            po2 = [0, yt, 0]
            objidlist = acad.GetSelectCornerCrossIdList(pt1, pt2)
            for objid in objidlist: acad.TransMove(objid, po1, po2)


@acad.decorator_command
def llalign_distance_region3():
    zhu_ui_llalign_distance()
    objidlist = acad.SSGetIdList(string="请选择对齐对象: ")  
    if objidlist == None: return
    with acad.transaction() as trans:
        result = acad.TransAutoFindRegionRectList(objidlist)[0:3]
        [pt1, pt2], [po1, po2], [pd1, pd2] = result