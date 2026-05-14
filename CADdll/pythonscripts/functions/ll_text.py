
import acad
import academit



def 命令(): 
    academit.添加命令("lltext-songti", lltext_songti)
    academit.添加命令("lltext-replace", lltext_replace)
    academit.添加命令("lltext-count-point-x-for", lltext_count_point_x_for)
    academit.添加命令("lltext-count-point-y-for", lltext_count_point_y_for)
    academit.添加命令("lltext-count-rec-for", lltext_count_rec_for)   
    academit.添加命令("lltext-count-frence-for", lltext_count_frence_for) 
    academit.添加命令("lltext-convert-from-attribute",lltext_convert_from_attribute)


zhu_count = 1
zhu_text_pre = ""
zhu_text_post = ""
zhu_text_size = 50
def zhu_uitext_count():
    global zhu_count, zhu_text_size, zhu_text_pre, zhu_text_post
    count = acad.GetInt(zhu_count, "请输入计数起点:")
    pre = acad.GetString(zhu_text_pre, "请输入前置文本")
    post = acad.GetString(zhu_text_post, "请输入后置文本")
    size = acad.GetInt(zhu_text_size, "请输入文字大小:")
    if count == None: count = 1
    if pre == None or pre == "0": pre = ""
    if post == None or post == "0": post = ""
    if size == None: size = 50
    zhu_count = count
    zhu_text_pre = pre
    zhu_text_post = post
    zhu_text_size = size


@acad.decorator_command
def lltext_songti():
    acad.ChangeStandardFontStyle("宋体")

    
@acad.decorator_command_undo
def lltext_count_point_x_for():
    zhu_uitext_count()
    global zhu_count
    pt0 = acad.GetPoint()
    if pt0 == None: return
    x0, y0, z0 = pt0
    char = f"{zhu_text_pre}{zhu_count}{zhu_text_post}"
    acad.CommandAddText(pt0, char, zhu_text_size)
    while True:
        pt1 = acad.GetPoint()
        zhu_count += 1
        if pt1 == None: break
        x1, y1, z1 = pt1
        pt1 = [x1, y0, z1]
        char = f"{zhu_text_pre}{zhu_count}{zhu_text_post}"
        acad.CommandAddText(pt1, char, zhu_text_size)

@acad.decorator_command_undo
def lltext_count_point_y_for():
    zhu_uitext_count()
    global zhu_count
    pt0 = acad.GetPoint()
    if pt0 == None: return
    x0, y0, z0 = pt0
    char = f"{zhu_text_pre}{zhu_count}{zhu_text_post}"
    acad.CommandAddText(pt0, char, zhu_text_size)
    while True:
        pt1 = acad.GetPoint()
        zhu_count += 1
        if pt1 == None: break
        x1, y1, z1 = pt1
        pt1 = [x0, y1, z1]
        char = f"{zhu_text_pre}{zhu_count}{zhu_text_post}"
        acad.CommandAddText(pt1, char, zhu_text_size)


@acad.decorator_command
def lltext_count_rec_for():
    zhu_uitext_count()
    global zhu_count
    objidlist = acad.SSGetIdList([[0, "LWPOLYLINE"]])
    with acad.transaction() as trans:
        for objid in objidlist:
            pt1 = acad.GetEntityBoundCenterXY0(objid)
            zhu_count += 1
            char = f"{zhu_text_pre}{zhu_count}{zhu_text_post}"
            acad.AddText(pt1, char, zhu_text_size)
        


@acad.decorator_command
def lltext_count_frence_for():
    zhu_uitext_count()
    pt1 = acad.GetPoint()
    pt2 = acad.GetPoint(base_point=pt1)
    ss1 = acad.GetSelectFence(pt1, pt2)
    objidlist = ss1.GetObjectIds()
    with acad.transaction() as trans:
        for i, objid in enumerate(objidlist):
            pt1 = acad.GetEntityBoundCenterXY0(objid)
            char = f"{zhu_text_pre}{i+1}{zhu_text_post}"
            acad.AddText(pt1, char, zhu_text_size)


@acad.decorator_command
def lltext_replace():
    char1 = acad.GetString("-金属板", "请输入要寻找的文字对象:")
    char2 = acad.GetString("", "请输入用于替换的文字:")
    if char2 == None: char2 = ""
    objidlist = acad.SSGetIdList([[0, "TEXT"]]) # [0, "MTEXT"], mtext not has TextString
    with acad.transaction() as trans:
        for objid in objidlist:
            objref = acad.TransObjectForWrite(objid)
            string = objref.TextString
            if char1 in string:
                string = string.replace(char1, char2)
                objref.TextString = string


@acad.decorator_command
def lltext_convert_from_attribute():
    objidlist = acad.SSGetIdList([[0, "ATTDEF"]])
    with acad.transaction() as trans:
        for objid in objidlist:
            objref = acad.TransObjectForWrite(objid)
            string = objref.Tag
            size = objref.Height
            pt1 = [objref.Position.X, objref.Position.Y, objref.Position.Z]
            text = acad.AddText(pt1, string, size)
            objref.Erase() # Need OpenForWrite
