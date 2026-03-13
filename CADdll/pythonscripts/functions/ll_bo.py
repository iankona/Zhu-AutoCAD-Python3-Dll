  
import acad
import academit

def 命令(): 
    academit.添加命令("llbo-for", llbo_for)
    academit.添加命令("llbo-color-index-set", llbo_color_index_set)
    academit.添加命令("llbo-offset-n-for", llbo_offset_n_for)
    academit.添加命令("llbo-offset-w-for", llbo_offset_w_for)
    academit.添加命令("llbo-add-rect-w-for", llbo_add_rect_w_for)



@acad.decorator_command
def llbo_color_index_set():
    index = acad.GetInt(0, string="请输入颜色索引: ")
    if index == None: return
    acad.CountColorIndexSet(index)



@acad.decorator_command
def llbo_for():
    color_index = acad.CountColorIndex()
    while True:
        pt1 = acad.GetPoint("请点击内部点: ")
        if pt1 == None: break
        objreflist = acad.ed.TraceBoundary(acad.ToPoint3d(pt1), True)
        with acad.transaction() as trans:
            for i, objref in enumerate(objreflist):
                objref.ColorIndex = color_index
                acad.AddDBObject(objref)

zhu_offset_n = 10
@acad.decorator_command
def llbo_offset_n_for():
    global zhu_offset_n
    length = acad.GetDouble(zhu_offset_n, "请输入内偏移大小:")
    if length != None: zhu_offset_n = length
    color_index = acad.CountColorIndex()
    while True:
        pt1 = acad.GetPoint("请点击内部点: ")
        if pt1 == None: break
        objreflist = acad.ed.TraceBoundary(acad.ToPoint3d(pt1), True)
        with acad.transaction() as trans:
            for objref1 in objreflist:
                collect2 = objref1.GetOffsetCurves(-zhu_offset_n)  
                for objref2 in collect2: 
                    objref2.ColorIndex = color_index
                    acad.AddDBObject(objref2)
                

zhu_offset_w = 20
@acad.decorator_command
def llbo_offset_w_for():
    global zhu_offset_w
    length = acad.GetDouble(zhu_offset_w, "请输入外偏移大小:")
    if length != None: zhu_offset_w = length
    color_index = acad.CountColorIndex()
    while True:
        pt1 = acad.GetPoint("请点击内部点: ")
        if pt1 == None: break
        objreflist = acad.ed.TraceBoundary(acad.ToPoint3d(pt1), True)
        with acad.transaction() as trans:
            for objref1 in objreflist:
                collect2 = objref1.GetOffsetCurves(zhu_offset_w)  
                for objref2 in collect2: 
                    objref2.ColorIndex = color_index
                    acad.AddDBObject(objref2)









zhu_offset_rect = 25
@acad.decorator_command
def llbo_add_rect_w_for():
    import ll_connet
    global zhu_offset_rect
    length = acad.GetDouble(zhu_offset_rect, "请输入外方管大小:")
    if length != None: zhu_offset_rect = length

    while True:
        pt1 = acad.GetPoint("请点击内部点: ")
        if pt1 == None: break
        objreflist = acad.ed.TraceBoundary(acad.ToPoint3d(pt1), True)
        with acad.transaction() as trans:
            objref1 = objreflist[0]
            ptnlist = []
            for i in range(objref1.NumberOfVertices):
                point = objref1.GetPoint3dAt(i)
                ptnlist.append([point.X, point.Y, point.Z])
            collect = objref1.GetOffsetCurves(zhu_offset_rect)  
            objref2 = collect[0] 
            ptwlist = []
            for i in range(objref2.NumberOfVertices):
                point = objref2.GetPoint3dAt(i)
                ptwlist.append([point.X, point.Y, point.Z])
            buflist = ll_connet.zhu_connet_rect_and_rect(ptnlist, ptwlist)
            for pt1, pt2 in buflist:
                x1, y1, z1 = pt1
                x2, y2, z2 = pt2
                po1 = [x1, y1, 0]
                po2 = [x1, y2, 0]
                po3 = [x2, y2, 0]
                po4 = [x2, y1, 0]
                acad.AddLWPolyLine([po1, po2, po3, po4, po1])

