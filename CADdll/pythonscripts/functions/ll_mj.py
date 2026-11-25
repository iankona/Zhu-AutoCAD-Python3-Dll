import clr

import acad
import academit

import System
from Autodesk.AutoCAD.DatabaseServices import RegAppTableRecord, ResultBuffer, TypedValue, DxfCode, ObjectId, Xrecord


def 命令(): 
    academit.添加命令("llpj", llpj)
    academit.添加命令("llmj1", llmj1)
    academit.添加命令("llmj", llmj)
    academit.添加命令("llmjbo", llmjbo)
    academit.添加命令("llbomj", llmjbo)
    academit.添加命令("llmj-xdata", llmj_xdata)
    academit.添加命令("llmj-xrecord", llmj_xrecord)



def ui_text_size():
    global zhu_text_size
    size = acad.GetInt(zhu_text_size, "请输入字体大小: ")
    if size > 0: zhu_text_size = size

zhu_text_size = 150
@acad.decorator_command
def llmj():
    ui_text_size()
    objid = acad.EntSel([[0, "LWPOLYLINE"]])
    pt1 = acad.GetPoint()
    if acad.IsNone(objid): return
    with acad.transaction() as trans:
        objref = acad.TransObjectForRead(objid)
        area = objref.Area / 1000000
        length = objref.Length / 1000
        text1 = f"面积:{area:.2f}m2"
        text2 = f"周长:{length:.2f}米"
        chars = "" + "\n" + text1 + "\n" + text2 + "\n" + "墙面积:m2" + "\n" + "墙顶面积:m2"+ "\n" + "踢脚线:米"
        acad.AddMText(pt1, chars, zhu_text_size)


zhu_text_size = 150
@acad.decorator_command
def llpj():
    ui_text_size()
    while True:
        objid = acad.EntSel([[0, "LWPOLYLINE"]])
        pt1 = acad.GetPoint()
        if acad.IsNone(objid): return
        with acad.transaction() as trans:
            objref = acad.TransObjectForRead(objid)
            # length = objref.Length / 1000
            # chars = f"总长:{length:.4f}米"
            chars = f"总长:{objref.Length:.2f}mm"
            # if chars.endswith("0mm"): chars = chars[0:-3]
            # if chars.endswith("00mm"): chars = chars[0:-5]
            acad.AddMText(pt1, chars, zhu_text_size)



@acad.decorator_command
def llmj1():
    ui_text_size()
    objid = acad.EntSel([[0, "LWPOLYLINE"]])
    pt1 = acad.GetPoint()
    if acad.IsNone(objid): return
    with acad.transaction() as trans:
        objref = acad.TransObjectForRead(objid)
        area = objref.Area / 1000000
        text = f"面积:{area:.2f}m2"
        acad.AddMText(pt1, text, zhu_text_size)





@acad.decorator_command
def llmjbo():
    ui_text_size()
    sumarea = 0
    while True:
        pt1 = acad.GetPoint("请点击内部点: ")
        if pt1 == None: break
        objreflist = acad.ed.TraceBoundary(acad.ToPoint3d(pt1), True)
        with acad.transaction() as trans:
            for i, objref in enumerate(objreflist):
                area = objref.Area / 1000000
                text = f"面积:{area:.2f}m2"
                sumarea += area
                acad.AddDBObject(objref, color_index=6)
                acad.AddMText(pt1, text, zhu_text_size, color_index=6)

    pt1 = acad.GetPoint("请点击位置: ")
    text = f"面积:{sumarea:.2f}m2"
    acad.AddMText(pt1, text, zhu_text_size, color_index=6)










class xdatadxfcode:
    flag = System.Int32(1001)
    text = System.Int32(1000)
    real = System.Int32(1040)
    vec3 = System.Int32(1010) # [0.0, 0.1, 0.3]
    int16 = System.Int32(1070)
    int32 = System.Int32(1071)
    







def AddRegTableRecord(name="ZhuRegAppTableRecord"):
    with acad.transaction() as trans:
        regapptable = acad.TransObjectForWrite(acad.db.RegAppTableId)
        if regapptable.Has(name): return None
        record = RegAppTableRecord()
        record.Name = name
        regapptable.Add(record)
        trans.AddNewlyCreatedDBObject(record, True)
        return None # xdata自动附加到了record里

# def AddXData(name="id"):
#     with acad.transaction() as trans:
#         regapptable = acad.TransObjectForRead(acad.db.RegAppTableId)
#         if regapptable.Has(name):






# XData支持多种数据类型，常用的包括：
# 1001: 应用程序名称（必须首先出现）
# 1000: ASCII字符串
# 1040: 实数
# 1070: 16位整数
# 1071: 32位整数


@acad.decorator_command
def llmj_xdata():
    # objid = acad.EntSel([[0, "LWPOLYLINE"]])
    # if acad.IsNone(objid): return
    objid = acad.EntSel()
    pt1, pt2 = acad.GetPoint2("", "", True)
    if pt1 == None: return
    with acad.transaction() as trans:
        mtext = acad.TransObjectForWrite(objid)
        char = mtext.Text
        AddRegTableRecord(name="ZhuRegAppTableRecord")
        resultbuf = ResultBuffer()
        resultbuf.Add(TypedValue(xdatadxfcode.flag, "ZhuRegAppTableRecord")) # 必须首先出现，固定字段
        resultbuf.Add(TypedValue(xdatadxfcode.text, "鹅鹅鹅"))
        resultbuf.Add(TypedValue(xdatadxfcode.real, 20.0))
        resultbuf.Add(TypedValue(xdatadxfcode.int32, 5))
        mtext.XData = resultbuf
        mtext.Contents = mtext.Contents + "\n我的扩展数据"
        





@acad.decorator_command
def llmj_xrecord():
    # objid = acad.EntSel([[0, "LWPOLYLINE"]])
    # if acad.IsNone(objid): return
    objid = acad.EntSel()
    with acad.transaction() as trans:
        mtext = acad.TransObjectForWrite(objid)
        if mtext.ExtensionDictionary == ObjectId.Null: mtext.CreateExtensionDictionary()
        xdict = acad.TransObjectForWrite(mtext.ExtensionDictionary)
        xrecord = Xrecord()
        AddRegTableRecord(name="ZhuRegAppTableRecord")
        resultbuf = ResultBuffer()
        resultbuf.Add(TypedValue(xdatadxfcode.flag, "ZhuRegAppTableRecord")) # 必须首先出现，固定字段
        resultbuf.Add(TypedValue(xdatadxfcode.text, "鹅鹅鹅"))
        resultbuf.Add(TypedValue(xdatadxfcode.real, 20.0))
        resultbuf.Add(TypedValue(xdatadxfcode.int32, 5))
        xrecord.Data = resultbuf
        xdict.SetAt("变量名", xrecord)
        trans.AddNewlyCreatedDBObject(xrecord, True)








# @acad.decorator_command
# def llcalc():
#     global zhu_text_size
#     string = ""
#     while True:
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



# @acad.decorator_command
# def llcalc():
#     global zhu_text_size
#     string = ""
#     while True:
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


