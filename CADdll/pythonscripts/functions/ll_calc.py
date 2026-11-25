import clr

import acad
import academit

import System

def 命令(): 
    academit.添加命令("llcalc", llcalc)
    academit.添加命令("llcalc-wall-area1", llcalc_wall_area1)
    academit.添加命令("llcalc-wall-area2", llcalc_wall_area2)
    academit.添加命令("llcalc-window-for", llcalc_window_for)

zhu_ceng_height = 2450
def ui_ceng_height():
    global zhu_ceng_height
    height= acad.GetInt(zhu_ceng_height, "请输入层高:")
    if height > 0: zhu_ceng_height = height

zhu_text_size = 50
def ui_text_size():
    global zhu_text_size
    size = acad.GetInt(zhu_text_size, "请输入字体大小:")
    if size > 0: zhu_text_size = size



@acad.decorator_command
def llcalc_wall_area1():
    ui_ceng_height()
    ui_text_size()
    objid = acad.EntSel([[0, "LWPOLYLINE"]])
    if acad.IsNone(objid): return

    dislist = []
    while True:
        pt1, pt2 = acad.GetPoint2("请选择第1点:", "请选择第2点:")
        if pt1 == None: break 
        distance = acad.Distance(pt1, pt2) / 1000
        dislist.append(distance)



    sumdis = 0
    for dis in dislist: sumdis += dis

    with acad.transaction() as trans:
        objref = acad.TransObjectForRead(objid)
        area = objref.Area / 1000000
        length = objref.Length / 1000
        height = zhu_ceng_height / 1000
        wall_area = length * height - sumdis
        sum_area = area + wall_area
        down_length = length - sumdis

        buft1 = ""
        for dis in dislist: buft1 += f"-{dis:.2f}"
        calc1 = f"={length:.2f}*{height:.2f}{buft1} = {wall_area:.2f}\n" 
        calc2 = f"={area:.2f}+{wall_area:.2f} = {sum_area:.2f}\n" 
        calc3 = f"={length:.2f}{buft1} = {down_length:.2f}\n\n" 
        text1 = f"面积:{area:.2f}m2\n"
        text2 = f"周长:{length:.2f}米\n"
        text3 = f"墙面积:{wall_area:.2f}m2\n"
        text4 = f"墙顶面积:{sum_area:.2f}m2\n"
        text5 = f"踢脚线:{down_length:.2f}米\n"
        string = calc1 + calc2 + calc3 + text1 + text2 + text3 + text4 + text5

        pt1 = acad.GetPoint("请点击放置位置:")
        if pt1 != None: acad.AddMText(pt1, string, zhu_text_size)

        

@acad.decorator_command
def llcalc_wall_area2():
    ui_ceng_height()
    ui_text_size()
    objid = acad.EntSel([[0, "LWPOLYLINE"]])
    if acad.IsNone(objid): return

    dislist = []
    while True:
        pt1, pt2= acad.GetPoint2("请选择第1点:", "请选择第2点:")
        if pt1 == None: break
        objidlist = acad.GetSelectCornerCrossIdList(pt1, pt2, [[-4, "<OR"],[0, "TEXT"],[0, "MTEXT"], [-4, "OR>"]])
        if objidlist == []:
            distance = acad.Distance(pt1, pt2) / 1000
            dislist.append(distance)
        else:
            text = get_text_string_from_id(objidlist[0])
            if text == None: break
            area = text_to_windown_area(text)
            dislist.append(area)

    sumdis = 0
    for dis in dislist: sumdis += dis

    with acad.transaction() as trans:
        objref = acad.TransObjectForRead(objid)
        area = objref.Area / 1000000
        length = objref.Length / 1000
        height = zhu_ceng_height / 1000
        wall_area = length * height - sumdis
        sum_area = area + wall_area
        down_length = length - sumdis

        buft1 = ""
        for dis in dislist: buft1 += f"-{dis:.2f}"
        calc1 = f"={length:.2f}*{height:.2f}{buft1} = {wall_area:.2f}\n" 
        calc2 = f"={area:.2f}+{wall_area:.2f} = {sum_area:.2f}\n" 
        calc3 = f"={length:.2f}{buft1} = {down_length:.2f}\n\n" 
        text1 = f"面积:{area:.2f}m2\n"
        text2 = f"周长:{length:.2f}米\n"
        text3 = f"墙面积:{wall_area:.2f}m2\n"
        text4 = f"墙顶面积:{sum_area:.2f}m2\n"
        text5 = f"踢脚线:{down_length:.2f}米\n"
        string = calc1 + calc2 + calc3 + text1 + text2 + text3 + text4 + text5

        pt1 = acad.GetPoint("请点击放置位置:")
        if pt1 != None: acad.AddMText(pt1, string, zhu_text_size)

        


@acad.decorator_command
def llcalc():
    string = ""
    while True:
        char = acad.GetString("请输入表达式:")
        if char == None: break
        string += char
        
    string = string.replace(" ", "")
    string = string.replace("=", "")
    string = string.replace("\n", "")
    string = string.replace("请输入表达式:", "")
    
    if string != "":   
        result = eval(string)
        string = f"= {string} = {result:.2f}"
        acad.Prompt(string)

    ui_text_size()
    pt1  = acad.GetPoint("请点击放置位置: ")
    with acad.transaction() as trans:
        acad.AddText(pt1, string, zhu_text_size)
    # print(string)


def get_text_string():
    objid = acad.EntSel([[-4, "<OR"],[0, "TEXT"],[0, "MTEXT"], [-4, "OR>"]], "请选择文本:")
    string = None
    with acad.transaction():
        objref = acad.TransObjectForRead(objid)
        match str(objref): 
            case "Autodesk.AutoCAD.DatabaseServices.DBText": string = objref.TextString
            case "Autodesk.AutoCAD.DatabaseServices.MText": string = objref.Text # 只读对像,无格式 # objref.Contents是文本写入对象，有格式 # 'SimSun|b0|i0|c134|p2;窗下:}525mm\\P{\\fSimSun|b0|i0|c134|p2;窗高:1235m'
    return string


def get_text_string_from_id(objid):
    string = None
    with acad.transaction():
        objref = acad.TransObjectForRead(objid)
        match str(objref): 
            case "Autodesk.AutoCAD.DatabaseServices.DBText": string = objref.TextString
            case "Autodesk.AutoCAD.DatabaseServices.MText": string = objref.Text # 只读对像,无格式 # objref.Contents是文本写入对象，有格式 # 'SimSun|b0|i0|c134|p2;窗下:}525mm\\P{\\fSimSun|b0|i0|c134|p2;窗高:1235m'
    return string


def text_to_height(string):
    string = string.replace("\r", "")
    lines = string.split("\n")
    a = None
    b = None
    for line in lines:
        if "窗下" in line: a = line
        if "窗高" in line: b = line
    if a == None or b == None: return None
    a = a.replace(" ", "")
    b = b.replace(" ", "")
    print(a.encode('utf-8'), b.encode('utf-8'))
    # b'\xe7\xaa\x97\xe4\xb8\x8b:525mm\r' b'\xe7\xaa\x97\xe9\xab\x98:1235mm'
    a = a[3:-2]
    b = b[3:-2]
    return float(a), float(b)


# def text_to_windown_area(string):
#     string = string.replace("\r", "")
#     lines = string.split("\n")
#     a = None
#     for line in lines:
#         if "窗面积" in line: a = line
#     if a == None: return None
#     a = a.replace(" ", "")
#     print(a.encode('utf-8'))
#     a = a[4:-2]
#     return float(a)

def text_to_windown_area(string):
    string = string.replace("\r", "")
    lines = string.split("窗")
    a = None
    for line in lines:
        if "面积" in line: a = line
    if a == None: return None
    a = a.replace(" ", "")
    print(a.encode('utf-8'))
    a = a[3:-2]
    return float(a)


@acad.decorator_command
def llcalc_window_for():
    ui_text_size()
    while True:
        text = get_text_string()
        if text == None: return
        height1, height2 = text_to_height(text) # 窗下，窗高
        height2 = height2/1000
        height_half = height2/2

        pt1, pt2 = acad.GetPoint2("请选择窗户第1点:", "请选择窗户第2点:")
        if pt1 == None: return 
        length = acad.Distance(pt1, pt2)/1000
        area = length*height_half

        string1 = f"= {length:.2f} * {height2:.2f} / 2 = {area:.2f}\r\n"
        string2 = f"窗面积:{area:.2f}m2"
        string = string1 + string2
        pt1  = acad.GetPoint("请点击放置位置: ")
        with acad.transaction() as trans:
            acad.AddMText(pt1, string, zhu_text_size)






# @acad.decorator_command
# def llcalc_window_for():
#     global zhu_text_size
#     string = ""
#     while True:
#         objid = acad.EntSel([0, "LWPOLYLINE"])
#         distance = acad.GetDistance(0)
#         with acad.transaction() as trans:
#             mtext = acad.TransObjectForWrite(objid)
#             char = mtext.Text

    



#         char = acad.GetString("请输入表达式: ")
#         if char == None: break
#         string += char
#     size = acad.GetInt(zhu_text_size, "请输入字体大小: ")
#     if size > 0: zhu_text_size = size
#     pt1  = acad.GetPoint("请点击放置位置: ")
#     string = string.replace("=", "")
#     string = string.replace("\n", "")
#     if string != "":   
#         result = eval(string)
#         string = f"= {string} = {result}"
#         acad.Prompt(string)
#         with acad.transaction() as trans:
#             acad.AddText(pt1, string, zhu_text_size)
#     # print(string)


