  
import acad
import academit
import System


def 命令(): 
    academit.添加命令("sb", sb)
    academit.添加命令("sbx", sbx)
    academit.添加命令("sby", sby)
    academit.添加命令("sba", sba)


    academit.添加命令("m-for", m_for)
    academit.添加命令("mx", mx)
    academit.添加命令("my", my)

    academit.添加命令("cox", cox)
    academit.添加命令("coy", coy)

    academit.添加命令("angle", angle)
    academit.添加命令("roa", roa)
    academit.添加命令("roar", roar)
    academit.添加命令("roar-invert", roar_invert)

    academit.添加命令("roc", roc)
    academit.添加命令("xlv", xlv)
    academit.添加命令("xlh", xlh)
    academit.添加命令("xlper", xlper)

    academit.添加命令("f0", f0)
    academit.添加命令("l0", lay0)
    academit.添加命令("lay0", lay0)

    academit.添加命令("zooma", zooma)
    academit.添加命令("zoomo", zoomo)
    academit.添加命令("larc", larc)
    academit.添加命令("larc3", larc3)

    academit.添加命令("dset", dset)

    academit.添加命令("midm", midm)
    academit.添加命令("midm3", midm3)
    academit.添加命令("midm3x", midm3x)
    academit.添加命令("midm3y", midm3y)

    academit.添加命令("midm4", midm4)
    academit.添加命令("midme", midme)

    academit.添加命令("midl", midl)
    academit.添加命令("midxl", midxl)

    academit.添加命令("midl1", midl1)
    academit.添加命令("midl2", midl2)
    academit.添加命令("linl2", linl2)
    academit.添加命令("linc2", linc2)


    academit.添加命令("chuang", chuang)
    academit.添加命令("chuang2", chuang2)
    academit.添加命令("ungroupdict", ungroupdict)
    academit.添加命令("rec50", rect50)
    academit.添加命令("rec100", rect100)
    academit.添加命令("rec300", rect300)
    academit.添加命令("rec400", rect400)
    academit.添加命令("rec550", rect550)
    academit.添加命令("rec600", rect600)        
    academit.添加命令("recx", rectdis)
    academit.添加命令("recdis", rectdis)
    academit.添加命令("recdislx", rectdislx2)
    academit.添加命令("dislx1", rectdislx1)
    academit.添加命令("dislx2", rectdislx2)

    academit.添加命令("hsolid", hsolid)
    academit.添加命令("lladt2", lladt2)

    academit.添加命令("brk2", break2)
    academit.添加命令("brk3", break3)
    academit.添加命令("brk4", break4)
    academit.添加命令("brk5", break5)
    academit.添加命令("brkx", breakx)

    academit.添加命令("extz", extz)


    academit.添加命令("copy-viewport-direct", copy_viewport_direct)
    academit.添加命令("mvoff", mvoff)
    academit.添加命令("mvon", mvon)





@acad.decorator_command
def mvoff():
    ss1 = acad.SSGet([[0, "VIEWPORT"]])
    acad.Command(["MVIEW", "OFF", ss1, ""])


@acad.decorator_command
def mvon():
    ss1 = acad.SSGet([[0, "VIEWPORT"]])
    acad.Command(["MVIEW", "ON", ss1, ""])


@acad.decorator_command
def copy_viewport_direct():
    objidlist = acad.SSGetIdList([[0, "VIEWPORT"]])
    with acad.transaction() as trans:
        objreflist = acad.TransObjectIdListForWrite(objidlist)
        vp = objreflist[0]
        for objref in objreflist[1:]:
            objref.ViewCenter = vp.ViewCenter
            objref.ViewHeight = vp.ViewHeight
            objref.ViewDirection = vp.ViewDirection
            objref.ViewTarget = vp.ViewTarget






@acad.decorator_command
def lladt2():
    pt1, pt2 = acad.GetPoint2("选择门宽第1点:", "选择门宽第2点:", baseptflag=True)
    pt1, pt2 = acad.GetAttachNDirectPointList(pt1, pt2, 50)
    with acad.transaction() as trans:
        acad.AddLine(pt1, pt2)
    # po1, po2 = acad.GetPoint2("选择门厚第1点:", "选择门厚第2点:", baseptflag=True)   
    # distance = acad.Distance(po1, po2) 
    # acad.Command(["ADT2"])
    # acad.Command(["ADT2", acad.ToPoint3d(pt1), acad.ToPoint3d(pt2), acad.ToPoint3d(po1), acad.ToPoint3d(po2)])

@acad.decorator_command
def hsolid():
    while True:
        pt1 = acad.GetPoint("选择基点:")
        acad.Command(["-hatch", "p", "s", "co", System.Int32(8), acad.ToPoint3d(pt1), "", ""])



@acad.decorator_command
def roc():
    ss1 = acad.SSGet()
    pt1 = acad.GetPoint("选择基点:")
    acad.Command(["rotate", ss1, "", acad.ToPoint3d(pt1), "C"])
 
@acad.decorator_command
def xlv():
    acad.Command(["xline", "v"])

@acad.decorator_command
def xlh():
    acad.Command(["xline", "h"])



@acad.decorator_command
def xlper():
    pt1 = acad.GetPoint()
    acad.Command(["xline", "per", acad.ToPoint3d(pt1), acad.ToPoint3d(pt1),""])

@acad.decorator_command
def zooma():
    acad.Command(["zoom", "a"])

@acad.decorator_command
def zoomo():
    acad.Command(["zoom", "o"])

@acad.decorator_command
def f0():
    acad.Command(["f", "r", acad.System.Int32(0)])

@acad.decorator_command
def lay0():
    acad.Command(["-layer", "s", "0", ""])


@acad.decorator_command
def sb():
    acad.Command(["STRETCH", "b"])

@acad.decorator_command
def sbx():
    po1, po2 = acad.GetCorner2()
    if po1 == None: return
    pt1, pt2 = acad.GetPoint2("", "", True)
    if pt1 == None: return
    pt1, pt2 = acad.HorizontalPt1Pt2(pt1, pt2)
    ss1 = acad.GetSelectCornerCross(po1, po2)
    acad.Command(["STRETCH", "b",  acad.ToPoint3d(po1),  acad.ToPoint3d(po2), ss1, "",  acad.ToPoint3d(pt1),  acad.ToPoint3d(pt2)])

@acad.decorator_command
def sby():
    po1, po2 = acad.GetCorner2()
    if po1 == None: return
    pt1, pt2 = acad.GetPoint2("", "", True)
    if pt1 == None: return
    pt1, pt2 = acad.VerticalPt1Pt2(pt1, pt2)
    ss1 = acad.GetSelectCornerCross(po1, po2)
    acad.Command(["STRETCH", "b",  acad.ToPoint3d(po1),  acad.ToPoint3d(po2), ss1, "",  acad.ToPoint3d(pt1),  acad.ToPoint3d(pt2)])


@acad.decorator_command
def sba():
    po1, po2 = acad.GetCorner2()
    if po1 == None: return
    pt1, pt2 = acad.GetPoint2("", "", True)
    if pt1 == None: return
    axis = acad.AxisDirect(pt1, pt2)
    match axis:
        case "+X"|"-X": pt1, pt2 = acad.HorizontalPt1Pt2(pt1, pt2)
        case "+Y"|"-Y": pt1, pt2 = acad.VerticalPt1Pt2(pt1, pt2)
    ss1 = acad.GetSelectCornerCross(po1, po2)
    acad.Command(["STRETCH", "b",  acad.ToPoint3d(po1),  acad.ToPoint3d(po2), ss1, "",  acad.ToPoint3d(pt1),  acad.ToPoint3d(pt2)])


@acad.decorator_command
def sx():
    objidlist = acad.SSGetIdList()
    sx1 = acad.GetDouble(0.5)
    with acad.transaction() as trans:
        for objid in objidlist: acad.TransScale(objid, [0,0,0], sx1)

@acad.decorator_command
def sxre():
    objidlist = acad.SSGetIdList()
    pt1, pt2 = acad.GetPoint2()
    distance1 = acad.Distance(pt1, pt2)
    distance2 = acad.GetDouble(distance1, "请输入距离大小:")
    sx1 = distance2 / distance1
    with acad.transaction() as trans:
        for objid in objidlist: acad.TransScale(objid, [0,0,0], sx1)





@acad.decorator_command
def cox():
    while True:
        ss1 = acad.SSGet()
        if ss1 == None: return
        pt1, pt2 = acad.GetPoint2()
        if pt1 == None: return
        x1, y1, z1 = pt1
        x2, y2, z2 = pt2
        acad.Command(["copy", ss1, "", acad.ToPoint3d([x1, 0, 0]), acad.ToPoint3d([x2, 0, 0])])


@acad.decorator_command
def coy():
    while True:
        ss1 = acad.SSGet()
        if ss1 == None: return
        pt1, pt2 = acad.GetPoint2()
        if pt1 == None: return
        x1, y1, z1 = pt1
        x2, y2, z2 = pt2
        acad.Command(["copy", ss1, "", acad.ToPoint3d([0, y1, 0]), acad.ToPoint3d([0, y2, 0])])






@acad.decorator_command
def m_for():
    pt1, pt2 = acad.GetPoint2("", "", True)
    if pt1 == None: return
    for i in range(50):
        ss1 = acad.SSGet()
        if ss1 == None: break
        x1, y1, z1 = pt1
        x2, y2, z2 = pt2
        acad.Command(["move", ss1, "", acad.ToPoint3d([x1, y1, 0]), acad.ToPoint3d([x2, y2, 0])])



@acad.decorator_command
def mx():
    ss1 = acad.SSGet()
    pt1, pt2 = acad.GetPoint2("", "", True)
    if pt1 == None: return
    x1, y1, z1 = pt1
    x2, y2, z2 = pt2
    acad.Command(["move", ss1, "", acad.ToPoint3d([x1, 0, 0]), acad.ToPoint3d([x2, 0, 0])])




@acad.decorator_command
def my():
    ss1 = acad.SSGet()
    pt1, pt2 = acad.GetPoint2("", "", True)
    if pt1 == None: return
    x1, y1, z1 = pt1
    x2, y2, z2 = pt2
    acad.Command(["move", ss1, "", acad.ToPoint3d([0, y1, 0]), acad.ToPoint3d([0, y2, 0])])



@acad.decorator_command
def angle():
    pt1, pt2, pt3 = acad.GetPoint3()
    dr1 = acad.Direct(pt2, pt1)
    dr2 = acad.Direct(pt2, pt3)
    angle = acad.AngleFromDotDr1Dr2(dr1, dr2)
    acad.Prompt(angle)


@acad.decorator_command
def roa():
    ss1 = acad.SSGet()
    if ss1 == None: return
    pt1, pt2 = acad.GetPoint2(baseptflag=True)
    if pt1 == None: return
    angle = acad.GetDouble(-90)
    acad.Command(["rotate3d", ss1, "", acad.ToPoint3d(pt1), acad.ToPoint3d(pt2), angle])

@acad.decorator_command
def roar():
    ss1 = acad.SSGet()
    if ss1 == None: return
    pt1, pt2 = acad.GetPoint2("请选择轴第1点", "请选择轴第2点", baseptflag=True)
    if pt1 == None: return
    po1, po2, po3 = acad.GetPoint3()
    dr1 = acad.Direct(po2, po1)
    dr2 = acad.Direct(po2, po3)
    angle = acad.AngleFromDotDr1Dr2(dr1, dr2)
    angle = acad.GetDouble(angle)
    acad.Command(["rotate3d", ss1, "", acad.ToPoint3d(pt1), acad.ToPoint3d(pt2), angle])


@acad.decorator_command
def roar_invert():
    ss1 = acad.SSGet()
    if ss1 == None: return
    pt1, pt2 = acad.GetPoint2("请选择轴第1点", "请选择轴第2点", baseptflag=True)
    if pt1 == None: return
    po1, po2, po3 = acad.GetPoint3()
    dr1 = acad.Direct(po2, po1)
    dr2 = acad.Direct(po2, po3)
    angle = acad.AngleFromDotDr1Dr2(dr1, dr2)
    angle = acad.GetDouble(180-angle)
    acad.Command(["rotate3d", ss1, "", acad.ToPoint3d(pt1), acad.ToPoint3d(pt2), angle])


@acad.decorator_command
def dset():
    objid = acad.db.Dimstyle
    with acad.transaction() as trans:
        record = acad.TransObjectForRead(objid)
        dimscale = record.Dimscale
    numb = int(dimscale)
    biaohao = acad.GetInt(numb, "请输入标注号:")
    acad.SetCurrentDimStyle(f"副本{biaohao} ISO-25")

@acad.decorator_command
def ungroupdict():
    with acad.transaction() as trans:
        groupdict = acad.TransObjectForWrite(acad.db.GroupDictionaryId)
        for ide in groupdict: # DictionaryEntry 'DictionaryEntry' object has no attribute 'Erase'
            partgroup = acad.TransObjectForWrite(ide.Value) # Erase 必须配合 Write 对象使用
            partgroup.Erase(True)
            # partgroup = acad.TransObjectForRead(ide.Value)
            # partgroup.UpgradeOpen()
            # partgroup.Erase(True)
            # partgroup.DowngradeOpen()


@acad.decorator_command
def midm():
    while True:
        ss1 = acad.SSGet()
        pt1, pt2 = acad.GetPoint2("", "", True)
        if pt1 == None: return
        pt1, pt2 = acad.PtXYZtoPtXY0([pt1, pt2]) 
        mid = acad.MidPt1Pt2(pt1, pt2)
        acad.Command(["move", ss1, "", acad.ToPoint3d(pt1), acad.ToPoint3d(mid)])



@acad.decorator_command
def midm3():
    while True:
        ss1 = acad.SSGet()
        pt1 = acad.GetPoint("请选择基点:")
        pt2, pt3 = acad.GetPoint2("请选择第1点:", "请选择第2点:", baseptflag=True)
        if pt1 == None: return
        pt1, pt2, pt3 = acad.PtXYZtoPtXY0([pt1, pt2, pt3]) 
        mid = acad.MidPt1Pt2(pt2, pt3)
        acad.Command(["move", ss1, "", acad.ToPoint3d(pt1), acad.ToPoint3d(mid)])


@acad.decorator_command
def midm3x():
    ss1 = acad.SSGet()
    pt1 = acad.GetPoint("请选择基点:")
    pt2, pt3 = acad.GetPoint2("请选择第1点:", "请选择第2点:", baseptflag=True)
    if pt1 == None: return
    pt1, pt2, pt3 = acad.PtXYZtoPtXY0([pt1, pt2, pt3]) 
    mid = acad.MidPt1Pt2(pt2, pt3)
    x1, y1, z1 = pt1
    x2, y2, z2 = mid
    acad.Command(["move", ss1, "", acad.ToPoint3d([x1,0,0]), acad.ToPoint3d([x2,0,0])])

@acad.decorator_command
def midm3y():
    ss1 = acad.SSGet()
    pt1 = acad.GetPoint("请选择基点:")
    pt2, pt3 = acad.GetPoint2("请选择第1点:", "请选择第2点:", baseptflag=True)
    if pt1 == None: return
    pt1, pt2, pt3 = acad.PtXYZtoPtXY0([pt1, pt2, pt3]) 
    mid = acad.MidPt1Pt2(pt2, pt3)
    x1, y1, z1 = pt1
    x2, y2, z2 = mid
    acad.Command(["move", ss1, "", acad.ToPoint3d([0,y1,0]), acad.ToPoint3d([0,y2,0])])


@acad.decorator_command
def midm4():
    while True:
        ss1 = acad.SSGet()
        pt1, pt2, pt3, pt4 = acad.GetPoint4()
        if pt1 == None: return
        pt1, pt2, pt3, pt4 = acad.PtXYZtoPtXY0([pt1, pt2, pt3, pt4]) 
        mid1 = acad.MidPt1Pt2(pt1, pt2)
        mid2 = acad.MidPt1Pt2(pt3, pt4)
        acad.Command(["move", ss1, "", acad.ToPoint3d(mid1), acad.ToPoint3d(mid2)])


@acad.decorator_command
def midme():
    while True:
        ss1 = acad.SSGet()
        pt1, pt2 = acad.GetPoint2()
        if pt1 == None: return
        objid = acad.EntSel()
        if acad.IsNone(objid): return
        pt1, pt2 = acad.PtXYZtoPtXY0([pt1, pt2]) 
        mid1 = acad.MidPt1Pt2(pt1, pt2)
        mid2 = acad.GetEntityBoundCenterXY0(objid)
        acad.Command(["move", ss1, "", acad.ToPoint3d(mid1), acad.ToPoint3d(mid2)])

@acad.decorator_command
def midl():
    pt1, pt2, pt3 = acad.GetPoint3("请选择第1个顶点: ", "请选择第2个顶点: ", "请点击方向顶点: ", True)
    if pt1 == None: return
    pt1, pt2, pt3 = acad.PtXYZtoPtXY0([pt1, pt2, pt3]) 
    with acad.transaction() as trans:
        dr1 = acad.GetPerDirectXY(pt1, pt2, pt3) 
        po1 = acad.MidPt1Pt2(pt1, pt2)
        po2 = acad.Vec3Add(po1, dr1)
        acad.AddLine(po1, po2)


@acad.decorator_command
def midd():
    pass
    # pt1, pt2, pt3 = acad.GetPoint3("请选择第1个顶点: ", "请选择第2个顶点: ", "请点击方向顶点: ", True)
    # if pt1 == None: return
    # with acad.transaction() as trans:
    #     dr1 = acad.GetPerDirectXY(pt1, pt2, pt3) 
    #     po1 = acad.MidPt1Pt2(pt1, pt2)
    #     po2 = acad.Vec3Add(po1, dr1)
    #     acad.AddLine(po1, po2)

@acad.decorator_command
def midxl():
    pt1, pt2, pt3 = acad.GetPoint3("请选择第1个顶点: ", "请选择第2个顶点: ", "请点击方向顶点: ", True)
    if pt1 == None: return
    pt1, pt2, pt3 = acad.PtXYZtoPtXY0([pt1, pt2, pt3]) 
    dr1 = acad.GetPerDirectXY(pt1, pt2, pt3) 
    po1 = acad.MidPt1Pt2(pt1, pt2)
    po2 = acad.Vec3Add(po1, dr1)
    acad.Command(["XLINE", acad.ToPoint3d(po1), acad.ToPoint3d(po2), ""])





@acad.decorator_command
def larc():
    pt1, pt2, pt3 = acad.GetPoint3("请选择第1个顶点: ", "请选择第2个顶点: ", "请点击方向顶点: ", True)
    if pt1 == None: return
    pt1, pt2, pt3 = acad.PtXYZtoPtXY0([pt1, pt2, pt3]) 
    with acad.transaction() as trans:
        distance = acad.Distance(pt1, pt2)
        dr1 = acad.GetPerDirectResetLengthXY(pt1, pt2, pt3, distance/9)
        mid = acad.MidPt1Pt2(pt1, pt2)
        pt3 = acad.Vec3Add(mid, dr1)
        acad.Command(["_arc", acad.ToPoint3d(pt1), acad.ToPoint3d(pt3), acad.ToPoint3d(pt2)])


@acad.decorator_command
def larc3():
    pt1, pt2, pt3, pt4 = acad.GetPoint4("请选择开关点位:", "请选择灯具点位第1个顶点:", "请选择灯具点位第2个顶点:", "请点击方向顶点:")
    if pt1 == None: return
    pt1, pt2, pt3 = acad.PtXYZtoPtXY0([pt1, pt2, pt3]) 
    with acad.transaction() as trans:
        pt2 = acad.MidPt1Pt2(pt2, pt3)
        pt3 = pt4
        distance = acad.Distance(pt1, pt2)
        dr1 = acad.GetPerDirectResetLengthXY(pt1, pt2, pt3, distance/9)
        mid = acad.MidPt1Pt2(pt1, pt2)
        pt3 = acad.Vec3Add(mid, dr1)
        acad.Command(["_arc", acad.ToPoint3d(pt1), acad.ToPoint3d(pt3), acad.ToPoint3d(pt2)])








@acad.decorator_command
def midl1():
    import ll_yx
    ll_yx.zhu_ui_llyx_offset()
    pt1, pt2, pt3 = acad.GetPoint3("请选择第1个顶点: ", "请选择第2个顶点: ", "请点击方向顶点: ", True)
    if pt1 == None: return
    pt1, pt2, pt3 = acad.PtXYZtoPtXY0([pt1, pt2, pt3]) 
    with acad.transaction() as trans:
        axis = acad.AxisDirect(pt1, pt2)
        match axis:
            case "+X"|"-X": pt1, pt2 = acad.HorizontalPt1Pt2(pt1, pt2)
            case "+Y"|"-Y": pt1, pt2 = acad.VerticalPt1Pt2(pt1, pt2)
        dr1 = acad.GetPerDirectXY(pt1, pt2, pt3) 
        po1 = acad.MidPt1Pt2(pt1, pt2)
        po2 = acad.Vec3Add(po1, dr1)
        acad.AddLine(po1, po2)




@acad.decorator_command
def midl2():
    import ll_yx
    ll_yx.zhu_ui_llyx_offset()
    pt1, pt2, pt3 = acad.GetPoint3("请选择第1个顶点: ", "请选择第2个顶点: ", "请点击方向顶点: ", True)
    if pt1 == None: return
    pt1, pt2, pt3 = acad.PtXYZtoPtXY0([pt1, pt2, pt3]) 
    with acad.transaction() as trans:
        axis = acad.AxisDirect(pt1, pt2)
        match axis:
            case "+X"|"-X": pt1, pt2 = acad.HorizontalPt1Pt2(pt1, pt2)
            case "+Y"|"-Y": pt1, pt2 = acad.VerticalPt1Pt2(pt1, pt2)
        dr1 = acad.GetPerDirectXY(pt1, pt2, pt3) 
        pt1 = acad.MidPt1Pt2(pt1, pt2)
        pt2 = acad.Vec3Add(pt1, dr1)
        distance = ll_yx.zhu_llyx_offset_length / 2
        dr1 = acad.GetPerDirectWithPerflagResetLengthXY(pt1, pt2, +1, distance)
        dr2 = acad.GetPerDirectWithPerflagResetLengthXY(pt1, pt2, -1, distance)
        po1 = acad.Vec3Add(pt1, dr1)
        po2 = acad.Vec3Add(pt2, dr1)
        pd1 = acad.Vec3Add(pt1, dr2)
        pd2 = acad.Vec3Add(pt2, dr2)
        acad.AddLine(po1, po2)
        acad.AddLine(pd1, pd2)


@acad.decorator_command
def linl2():
    import ll_yx
    ll_yx.zhu_ui_llyx_offset()
    pt1, pt2, pt3 = acad.GetPoint3("请选择第1个顶点: ", "请选择第2个顶点: ", "请点击方向顶点: ", True)
    if pt1 == None: return
    pt1, pt2, pt3 = acad.PtXYZtoPtXY0([pt1, pt2, pt3]) 
    with acad.transaction() as trans:
        axis = acad.AxisDirect(pt1, pt2)
        match axis:
            case "+X"|"-X": pt1, pt2 = acad.HorizontalPt1Pt2(pt1, pt2)
            case "+Y"|"-Y": pt1, pt2 = acad.VerticalPt1Pt2(pt1, pt2)
        dr1 = acad.GetPerDirectResetLengthXY(pt1, pt2, pt3, ll_yx.zhu_llyx_offset_length)
        po1 = acad.Vec3Add(pt1, dr1)
        po2 = acad.Vec3Add(pt2, dr1)
        acad.AddLine(pt1, pt2)
        acad.AddLine(po1, po2)




@acad.decorator_command
def linc2():
    import ll_yx
    ll_yx.zhu_ui_llyx_offset()
    pt1, pt2 = acad.GetPoint2("请选择第1个顶点: ", "请选择第2个顶点: ", True)
    if pt1 == None: return
    pt1, pt2 = acad.PtXYZtoPtXY0([pt1, pt2]) 
    with acad.transaction() as trans:
        axis = acad.AxisDirect(pt1, pt2)
        match axis:
            case "+X"|"-X": pt1, pt2 = acad.HorizontalPt1Pt2(pt1, pt2)
            case "+Y"|"-Y": pt1, pt2 = acad.VerticalPt1Pt2(pt1, pt2)
        distance = ll_yx.zhu_llyx_offset_length / 2
        dr1 = acad.GetPerDirectWithPerflagResetLengthXY(pt1, pt2, +1, distance)
        dr2 = acad.GetPerDirectWithPerflagResetLengthXY(pt1, pt2, -1, distance)
        po1 = acad.Vec3Add(pt1, dr1)
        po2 = acad.Vec3Add(pt2, dr1)
        pd1 = acad.Vec3Add(pt1, dr2)
        pd2 = acad.Vec3Add(pt2, dr2)
        acad.AddLine(po1, po2)
        acad.AddLine(pd1, pd2)

@acad.decorator_command
def chuang():
    while True:
        pt1, pt2 = acad.GetPoint2("请选择第1个顶点: ", "请选择第2个顶点: ", True)
        pt3 = acad.GetPoint("请选择第3个顶点: ", pt2)
        if pt1 == None: return
        if pt3 == None: return
        pt1, pt2, pt3 = acad.PtXYZtoPtXY0([pt1, pt2, pt3]) 
        distance = acad.Distance(pt2, pt3)/3
        dr1 = acad.Direct(pt2, pt3)
        dr1 = acad.Vec3ResetLength(dr1, distance)
        with acad.transaction() as trans:
            for i in range(2):
                po1 = acad.Vec3Add(pt1, dr1)
                po2 = acad.Vec3Add(pt2, dr1)
                acad.AddLine(po1, po2)
                pt1, pt2 = po1, po2


@acad.decorator_command
def chuang2():
    while True:
        pt1, pt2 = acad.GetPoint2("请选择第1个顶点: ", "请选择第2个顶点: ", True)
        pt3 = acad.GetPoint("请选择第3个顶点: ", pt2)
        if pt1 == None: return
        if pt3 == None: return
        pt1, pt2, pt3 = acad.PtXYZtoPtXY0([pt1, pt2, pt3]) 
        distance = acad.Distance(pt2, pt3)/3
        dr1 = acad.Direct(pt2, pt3)
        dr1 = acad.Vec3ResetLength(dr1, distance)
        with acad.transaction() as trans:
            po1 = acad.Vec3Add(pt1, dr1)
            po2 = acad.Vec3Add(pt2, dr1)
            acad.AddLine(po1, po2)
            po1 = acad.Vec3Add(po1, dr1)
            po2 = acad.Vec3Add(po2, dr1)
            acad.AddLine(po1, po2)
            po1 = acad.Vec3Add(po1, dr1)
            po2 = acad.Vec3Add(po2, dr1)
            acad.AddLine(pt1, po1)
            acad.AddLine(pt2, po2)



@acad.decorator_command
def meng():
    pass


@acad.decorator_command
def rect600():
    while True:
        pt1, pt2, pt3 = acad.GetPoint3("请选择第1个顶点: ", "请选择第2个顶点: ", "请点击方向顶点: ", baseptflag=True)
        if pt1 == None: break
        with acad.transaction() as trans:
            dr1 = acad.GetPerDirectResetLengthXY(pt1, pt2, pt3, 600) # 点在线上, WhichSideOfLineXY 会出现cannot access local variable 'result' where it is not associated with a value
            po1 = acad.Vec3Add(pt1, dr1)
            po2 = acad.Vec3Add(pt2, dr1)
            acad.AddMKPolyLine([pt1, pt2, po2, po1],color_index=4)

@acad.decorator_command
def rect550():
    while True:
        pt1, pt2, pt3 = acad.GetPoint3("请选择第1个顶点: ", "请选择第2个顶点: ", "请点击方向顶点: ", baseptflag=True)
        if pt1 == None: break
        with acad.transaction() as trans:
            dr1 = acad.GetPerDirectResetLengthXY(pt1, pt2, pt3, 600) # 点在线上, WhichSideOfLineXY 会出现cannot access local variable 'result' where it is not associated with a value
            po1 = acad.Vec3Add(pt1, dr1)
            po2 = acad.Vec3Add(pt2, dr1)
            acad.AddMKPolyLine([pt1, pt2, po2, po1],color_index=4)


@acad.decorator_command
def rect300():
    while True:
        pt1, pt2, pt3 = acad.GetPoint3("请选择第1个顶点: ", "请选择第2个顶点: ", "请点击方向顶点: ", baseptflag=True)
        if pt1 == None: break
        with acad.transaction() as trans:
            dr1 = acad.GetPerDirectResetLengthXY(pt1, pt2, pt3, 300) # 点在线上, WhichSideOfLineXY 会出现cannot access local variable 'result' where it is not associated with a value
            po1 = acad.Vec3Add(pt1, dr1)
            po2 = acad.Vec3Add(pt2, dr1)
            acad.AddMKPolyLine([pt1, pt2, po2, po1],color_index=4)

@acad.decorator_command
def rect400():
    while True:
        pt1, pt2, pt3 = acad.GetPoint3("请选择第1个顶点: ", "请选择第2个顶点: ", "请点击方向顶点: ", baseptflag=True)
        if pt1 == None: break
        with acad.transaction() as trans:
            dr1 = acad.GetPerDirectResetLengthXY(pt1, pt2, pt3, 400) # 点在线上, WhichSideOfLineXY 会出现cannot access local variable 'result' where it is not associated with a value
            po1 = acad.Vec3Add(pt1, dr1)
            po2 = acad.Vec3Add(pt2, dr1)
            acad.AddMKPolyLine([pt1, pt2, po2, po1], color_index=4)


@acad.decorator_command
def rect100():
    while True:
        pt1, pt2, pt3 = acad.GetPoint3("请选择第1个顶点: ", "请选择第2个顶点: ", "请点击方向顶点: ", baseptflag=True)
        if pt1 == None: break
        with acad.transaction() as trans:
            dr1 = acad.GetPerDirectResetLengthXY(pt1, pt2, pt3, 100) # 点在线上, WhichSideOfLineXY 会出现cannot access local variable 'result' where it is not associated with a value
            po1 = acad.Vec3Add(pt1, dr1)
            po2 = acad.Vec3Add(pt2, dr1)
            acad.AddMKPolyLine([pt1, pt2, po2, po1], color_index=4)



@acad.decorator_command
def rect50():
    while True:
        pt1, pt2, pt3 = acad.GetPoint3("请选择第1个顶点: ", "请选择第2个顶点: ", "请点击方向顶点: ", baseptflag=True)
        if pt1 == None: break
        with acad.transaction() as trans:
            dr1 = acad.GetPerDirectResetLengthXY(pt1, pt2, pt3, 50) # 点在线上, WhichSideOfLineXY 会出现cannot access local variable 'result' where it is not associated with a value
            po1 = acad.Vec3Add(pt1, dr1)
            po2 = acad.Vec3Add(pt2, dr1)
            acad.AddMKPolyLine([pt1, pt2, po2, po1], color_index=4)





zhu_distance = 600
zhu_subcount = 4

@acad.decorator_command
def rectdis():
    global zhu_distance
    while True:
        dis = acad.GetDistance(zhu_distance, "请输入矩形宽度: ")
        if dis != None: zhu_distance = dis
        pt1, pt2, pt3 = acad.GetPoint3("请选择第1个顶点: ", "请选择第2个顶点: ", "请点击方向顶点: ", baseptflag=True)
        if pt1 == None: break
        with acad.transaction() as trans:
            dr1 = acad.GetPerDirectResetLengthXY(pt1, pt2, pt3, dis) # 点在线上, WhichSideOfLineXY 会出现cannot access local variable 'result' where it is not associated with a value
            po1 = acad.Vec3Add(pt1, dr1)
            po2 = acad.Vec3Add(pt2, dr1)
            acad.AddMKPolyLine([pt1, pt2, po2, po1], color_index=4)


@acad.decorator_command
def rectdislx2():
    global zhu_distance, zhu_subcount
    while True:
        distance = acad.GetDistance(zhu_distance, "请输入矩形宽度: ")
        subcount = acad.GetInt(zhu_subcount, "请输入分隔数量: ")
        if distance != None: zhu_distance = distance
        if subcount != None: zhu_subcount = subcount
        pt1, pt2, pt3 = acad.GetPoint3("请选择第1个顶点: ", "请选择第2个顶点: ", "请点击方向顶点: ")
        if pt1 == None: break
        with acad.transaction() as trans:
            per = acad.GetPerDirectResetLengthXY(pt1, pt2, pt3, distance) # 点在线上, WhichSideOfLineXY 会出现cannot access local variable 'result' where it is not associated with a value
            po1 = acad.Vec3Add(pt1, per)
            po2 = acad.Vec3Add(pt2, per)
            acad.AddMKPolyLine([pt1, pt2, po2, po1],  color_index=4) # 青
            sublength = acad.Distance(pt1, pt2)/subcount
            dr1 = acad.DirectResetLength(pt1, pt2, sublength)
            ptlist1 = [pt1]
            ptlist2 = [po1]
            for i in range(subcount):
                pt1 = acad.Vec3Add(pt1, dr1)
                po1 = acad.Vec3Add(po1, dr1)
                ptlist1.append(pt1)
                ptlist2.append(po1)
            for pt1, po1 in zip(ptlist1[1:-1], ptlist2[1:-1]):
                acad.AddLine(pt1, po1, color_index=8) # 灰
            plinelist1 = []
            plinelist2 = []
            for i, [pt1, po1] in enumerate(zip(ptlist1, ptlist2)):
                if i%2 == 0:
                    plinelist1.append(pt1)
                    plinelist2.append(po1)
                else:
                    plinelist1.append(po1)
                    plinelist2.append(pt1)

            acad.AddLWPolyLine(plinelist1,  color_index=8) # 灰
            acad.AddLWPolyLine(plinelist2,  color_index=8) # 灰



@acad.decorator_command
def rectdislx1():
    global zhu_distance, zhu_subcount
    while True:
        distance = acad.GetDistance(zhu_distance, "请输入矩形宽度: ")
        subcount = acad.GetInt(zhu_subcount, "请输入分隔数量: ")
        if distance != None: zhu_distance = distance
        if subcount != None: zhu_subcount = subcount
        pt1, pt2, pt3 = acad.GetPoint3("请选择第1个顶点: ", "请选择第2个顶点: ", "请点击方向顶点: ")
        if pt1 == None: break
        with acad.transaction() as trans:
            per = acad.GetPerDirectResetLengthXY(pt1, pt2, pt3, distance) # 点在线上, WhichSideOfLineXY 会出现cannot access local variable 'result' where it is not associated with a value
            po1 = acad.Vec3Add(pt1, per)
            po2 = acad.Vec3Add(pt2, per)
            acad.AddMKPolyLine([pt1, pt2, po2, po1],  color_index=4) # 青
            sublength = acad.Distance(pt1, pt2)/subcount
            dr1 = acad.DirectResetLength(pt1, pt2, sublength)
            ptlist1 = [pt1]
            ptlist2 = [po1]
            for i in range(subcount):
                pt1 = acad.Vec3Add(pt1, dr1)
                po1 = acad.Vec3Add(po1, dr1)
                ptlist1.append(pt1)
                ptlist2.append(po1)
            for pt1, po1 in zip(ptlist1[1:-1], ptlist2[1:-1]):
                acad.AddLine(pt1, po1, color_index=8) # 灰
            plinelist1 = []
            plinelist2 = []
            for i, [pt1, po1] in enumerate(zip(ptlist1[0:-1], ptlist2[1:])):
                if i%2 == 0:
                    plinelist1.append(pt1)
                    plinelist2.append(po1)
                else:
                    plinelist1.append(po1)
                    plinelist2.append(pt1)

            acad.AddLWPolyLine(plinelist1,  color_index=8) # 灰
            acad.AddLWPolyLine(plinelist2,  color_index=8) # 灰



def __breakcurve(subcount=2):
    while True:
        pt0, objid = acad.EntSelEntity(string="请选择要打断的对象:")
        if acad.IsNone(objid): return
        with acad.transaction() as trans:
            objref = acad.TransObjectForWrite(objid)
            sublength = objref.Length/subcount
            sumlengthlist = []
            sumlength = 0
            for i in range(subcount-1):
                sumlength += sublength
                sumlengthlist.append(sumlength)

            paralist = []
            for sumlength in sumlengthlist:
                pa1 = objref.GetParameterAtDistance(sumlength)
                paralist.append(pa1)
            if paralist != []:
                paralist.sort()
                collection = acad.DoubleCollection()
                for para in paralist: 
                    collection.Add(para) 
                result = objref.GetSplitCurves(collection)
                for i, objrefsub in enumerate(result):
                    objrefsub.ColorIndex = i+1
                    acad.AddDBObject(objrefsub)
                objref.Erase()


@acad.decorator_command
def break2():
    __breakcurve(2)

@acad.decorator_command
def break3():
    __breakcurve(3)

@acad.decorator_command
def break4():
    __breakcurve(4)

@acad.decorator_command
def break5():
    __breakcurve(5)

@acad.decorator_command
def breakx():
    subcount = acad.GetInt(0, "请输入分段数量:")
    __breakcurve(subcount)



zhu_extend_height = 2600
def input_extend_z_hight():
    global zhu_extend_height
    height = acad.GetDouble(zhu_extend_height)
    if height == None: return
    zhu_extend_height = height
    

@acad.decorator_command
def extz():
    input_extend_z_hight()
    while True:
        ss1 = acad.EntSel()
        pt1 = acad.Point3d([0,0,0])
        pt2 = acad.Point3d([0,0,zhu_extend_height])
        acad.Command(["extrude", ss1, "",  "d", pt1, pt2])
        # acad.Prompt("\n")

@acad.decorator_command
def extzhalf():
    input_extend_z_hight()
    while True:
        ss1 = acad.EntSel()
        if ss1 == None: return
        pt1 = acad.Point3d([0,0,0])
        pt2 = acad.Point3d([0,0,zhu_extend_height/2])
        acad.Command(["extrude", ss1, "",  "d", pt1, pt2])
        # acad.Prompt("\n")