import math


import clr

import System

from Autodesk.AutoCAD.ApplicationServices import Application
from Autodesk.AutoCAD.EditorInput import PromptStatus, SelectionFilter, PromptSelectionOptions, SelectionSet, PromptIntegerOptions, PromptPointOptions, PromptDoubleOptions
# from Autodesk.AutoCAD.Runtime
from Autodesk.AutoCAD.DatabaseServices import Line, ObjectId, Transaction, OpenMode, BlockTable, BlockTableRecord, ObjectIdCollection, TypedValue, DxfCode, DwgVersion
from Autodesk.AutoCAD.DatabaseServices import Extents3d, Polyline, Line, Circle

from Autodesk.AutoCAD.Geometry import Point3d
# from Autodesk.AutoCAD.Internal.Utils import EntLast



# doc = Application.DocumentManager.MdiActiveDocument
# doc.Editor.WriteMessage("Hello, python, 鹅鹅鹅")
# ed = Application.DocumentManager.MdiActiveDocument.Editor


doc = None
ed = None
db = None
Command = None
CommandAsync = None


def GetActiveDocument():
    global doc, ed, db, Command, CommandAsync
    doc = Application.DocumentManager.MdiActiveDocument
    ed = doc.Editor
    db = doc.Database
    Command = doc.Editor.Command
    CommandAsync = doc.Editor.CommandAsync
    return doc, ed, db, Command, CommandAsync


def EntLast():
    # Command 和 CommandAsync 可能是多线程的特性，在大量线条Command生成时遇到了Command执行完但ed.SelectLast()获取不到entlast，ss1为None，引起CAD软件崩溃跳出
    # 需要获取线条绘制后的objid需要从Command绘图切换到db+lockdoc绘图
    # 后面测试
    pass
    # result = ed.SelectLast() # PromptSelectionResult # (OK,[((1375515652304),NonGraphical,0,)])
    # ss1 = result.Value # SelectionSet (((2361431560400),NonGraphical,0,))
    # if ss1 == None: return None
    # ids = ss1.GetObjectIds()
    # return ids[0]


def EntLastSet():
    result = ed.SelectLast()
    ss1 = result.Value
    return ss1


def SelectionSetFromID(objid:ObjectId):
    return SelectionSet.FromObjectIds([objid])


def ToPoint3d(pt0):
    try: x, y, z = pt0
    except: [x, y], z = pt0, 0
    return Point3d(x, y, z)


ll_old_osmode = 0
def GetOSMODE():
    global ll_old_osmode
    ll_old_osmode = Application.GetSystemVariable("OSMODE")
    Application.SetSystemVariable("OSMODE", System.Int32(0))


def SetOSMODE():
    Application.SetSystemVariable("OSMODE", System.Int32(ll_old_osmode))



def GetUndo():
    Command(["undo", "be"]), Prompt("\n")


def SetUndo():
    Command(["undo", "e"]), Prompt("\n")


def HappenErrorUndo():
    Command(["undo", "e"]), Prompt("\n")
    Command(["u"]), Prompt("\n")



# (setq ll_current_dimstyle nil)
# (defun ll-get-dimstyle()
# (setq ll_current_dimstyle (getvar "dimstyle"))
# )

# (defun ll-set-dimstyle()
# (command "dimstyle" "r" ll_current_dimstyle)
# )

# (defun ll-change-dimstyle(id / id dimname )
# ;(command "dimstyle" "r" "副本5 ISO-25")
# (setq dimname (strcat "副本" (itoa id) " ISO-25"))
# (command "dimstyle" "r" dimname)
# )




def Copy(objid:ObjectId|SelectionSet):
    Command(["copy", objid, "D", ""]), Prompt("\n")
    return EntLast()

def CopyMove(objid:ObjectId|SelectionSet, start_point, final_point):
    Command(["copy", objid, ToPoint3d(start_point), ToPoint3d(final_point),"E"]), Prompt("\n")
    return EntLast()

def Move(objid:ObjectId|SelectionSet, start_point, final_point):
    Command(["move", objid, ToPoint3d(start_point), ToPoint3d(final_point)]), Prompt("\n")
    return EntLast()

def OffSet(objid:ObjectId|SelectionSet, distance, directpt1=[], 图层=""): # directpt1 = [x, y]
    if 图层 == "当前":
        Command(["OFFSET", "L", "C", System.Double(distance), "E"]) # 当前图层
        Command(["OFFSET", "", objid, ToPoint3d(directpt1), "E"])
        Command(["OFFSET", "L", "S", "", "E"]) # 还原回原图层
    else: 
        Command(["OFFSET", System.Double(distance), objid, ToPoint3d(directpt1), "E"]) # 若是SelectionSet，则只会偏移第1个objid  
    Prompt("\n")
    return EntLast()

def Rotate(objid:ObjectId|SelectionSet, centert_point, angle):
    Command(["rotate", objid, ToPoint3d(centert_point), System.Double(angle)]), Prompt("\n")
    return EntLast()

def RotateCopy(objid:ObjectId|SelectionSet, centert_point, angle):
    Command(["rotate", objid, ToPoint3d(centert_point), "C", System.Double(angle)]), Prompt("\n")
    return EntLast()



def Erase(objid:ObjectId|SelectionSet):
    Command(["rotate", objid]), Prompt("\n")


def ChangeLayer(obj, layername):
    AddLayer(layername)
    obj.Layer = layername


def Normalized(x, y, z = 0):
    a = x**2 + y**2 + z**2
    distance = math.sqrt(a)
    xn, yn, zn = x/distance, y/distance, z/distance
    return [xn, yn, zn]

def Vec3ResetLength(dr1, length):
    x, y, z = Normalized(*dr1)
    x, y, z = x*length, y*length, z*length
    return [x, y, z]


def Vec2toVec3(pt0):
    try: x, y, z = pt0
    except: [x, y], z = pt0, 0
    return [x, y, z]


def Vec3Negative(pt0):
    x, y, z = Vec2toVec3(pt0)
    return [-x, -y, -z]    


def Vec3Add(pt1, pt2):
    x1, y1, z1 = Vec2toVec3(pt1)
    x2, y2, z2 = Vec2toVec3(pt2)
    return [x2+x1, y2+y1, z2+z1]


def Distance(pt1, pt2):
    x1, y1, z1 = Vec2toVec3(pt1)
    x2, y2, z2 = Vec2toVec3(pt2)
    a = (x2-x1)**2 + (y2-y1)**2 + (z2-z1)**2
    return math.sqrt(a)

def Direct(pt1, pt2):
    x1, y1, z1 = Vec2toVec3(pt1)
    x2, y2, z2 = Vec2toVec3(pt2)
    return [x2-x1, y2-y1, z2-z1]


def GetAttachNDirectPoints(pt1, pt2, length):
    dr1 = Direct(pt1, pt2)
    dr1 = Vec3ResetLength(dr1, length)
    dr2 = Direct(pt2, pt1)
    dr2 = Vec3ResetLength(dr2, length)    
    return [Vec3Add(pt1, dr1), Vec3Add(pt2, dr2)]


def GetAttachWDirectPoints(pt1, pt2, length):
    dr1 = Direct(pt2, pt1)
    dr1 = Vec3ResetLength(dr1, length)
    dr2 = Direct(pt1, pt2)
    dr2 = Vec3ResetLength(dr2, length)    
    return [Vec3Add(pt1, dr1), Vec3Add(pt2, dr2)]




def MidPt1Pt2(pt1, pt2):
    x1, y1, z1 = Vec2toVec3(pt1)
    x2, y2, z2 = Vec2toVec3(pt2)
    return [(x2+x1)/2, (y2+y1)/2, (z2+z1)/2]


def GetPerflagXY(pt1, pt2, pt3):
    return WhichSideOfLineXY(pt1, pt2, pt3)


def GetPerDirectWithPerflagXY(pt1, pt2, perflag):
    direct  = Direct(pt1, pt2)
    perdirect = DirectToPerDirectXY(direct, perflag)
    return perdirect


def GetPerDirectWithPerflagResetLengthXY(pt1, pt2, perflag, length):
    # if length == 0: raise ValueError("...长度为0...")
    direct  = Direct(pt1, pt2)
    perdirect = DirectToPerDirectXY(direct, perflag)
    x, y, z = Normalized(*perdirect)
    return [x*length, y*length, z*length]


def GetPerDirectXY(pt1, pt2, pt3):
    perflag = WhichSideOfLineXY(pt1, pt2, pt3)
    direct  = Direct(pt1, pt2)
    perdirect = DirectToPerDirectXY(direct, perflag)
    return perdirect





def GetPerDirectResetLengthXY(pt1, pt2, pt3, length):
    # if length == 0: raise ValueError("...长度为0...")
    perflag = WhichSideOfLineXY(pt1, pt2, pt3)
    direct  = Direct(pt1, pt2)
    perdirect = DirectToPerDirectXY(direct, perflag)
    x, y, z = Normalized(*perdirect)
    return [x*length, y*length, z*length]


def WhichSideOfLineXY(pt1, pt2, pt3):
    # 设线段端点为从 A(x1, y1)到 B(x2, y2), 线外一点P(x3，y3)，
    # 判断该点位于有向线 A→B 的那一侧。
    # a = (x2-x1, y2-y1)
    # b = (x3-x1, y3-y1)
    # a x b = | a | | b | sinφ (φ为两向量的夹角)
    # | a | | b |  ≠ 0 时，  a x b  决定点 P的位置
    # 所以  a x b  的 z 方向大小决定 P位置
    # (x2-x1)(y3-y1) – (y2-y1)(x3-x1)  >  0   左侧
    # (x2-x1)(y3-y1) – (y2-y1)(x3-x1)  <  0   右侧 
    # (x2-x1)(y3-y1) – (y2-y1)(x3-x1)  =  0   线段上
    x1, y1, z1 = Vec2toVec3(pt1)
    x2, y2, z2 = Vec2toVec3(pt2)
    x3, y3, z3 = Vec2toVec3(pt3)
    flag = (x2-x1) * (y3-y1) - (y2-y1) * (x3-x1)
    if flag > 0: flag =  1
    if flag < 0: flag = -1
    return flag


def DirectToPerDirectXY(dr0, perflag):
    x, y, z = dr0
    match perflag:
        case  1: return [-y,  x,  z]
        case -1: return [ y, -x,  z]
        case  _: raise ValueError("...点在线上...")


def AddPoint(pt1):
    Command(["point", ToPoint3d(pt1)]), Prompt("\n")
    return EntLast()


def AddText(pt1, size, string): # pt1 = [x, y, z]
    Command(["text", ToPoint3d(pt1), System.Int32(size), System.Int32(0), string]), Prompt("\n")
    return EntLast()


def AddLine(pt1, pt2): # pt1 = [x, y, z]
    Command(["LINE", ToPoint3d(pt1), ToPoint3d(pt2), ""]), Prompt("\n")
    return EntLast()


def AddPLine(*args): # 函数(pt1, pt2, pt3...)
    列表 = [ToPoint3d(pt1) for pt1 in args]
    Command(["PLINE"] + 列表 + [""]), Prompt("\n")
    return EntLast()

def AddRect(pt1, pt2):
    Command(["RECTANG", ToPoint3d(pt1), ToPoint3d(pt2), ""]), Prompt("\n")
    return EntLast()


def AddCircle(pt1, radius):
    AddCircleR(pt1, radius)

def AddCircleR(pt1, radius):
    Command(["CIRCLE", ToPoint3d(pt1), System.Double(radius), ""]), Prompt("\n")
    return EntLast()

def AddCircleD(pt1, diameter):
    Command(["CIRCLE", ToPoint3d(pt1), "D", System.Double(diameter), ""]), Prompt("\n")
    return EntLast()


def AddCircle2P(pt1, pt2):
    Command(["CIRCLE", "2P", ToPoint3d(pt1), ToPoint3d(pt2), ""]), Prompt("\n")
    return EntLast()

def AddCircle3P(pt1, pt2, pt3):
    Command(["CIRCLE", "3P", ToPoint3d(pt1), ToPoint3d(pt2), ToPoint3d(pt3), ""]), Prompt("\n")
    return EntLast()


def Prompt(string:str):
    ed.WriteMessage(str(string))


def EntSel(string: str=""):
    if string == "": string = "请选择对象: "
    result = ed.GetEntity(string)  # == AutoLisp entsel
    Prompt(f"(图元: {result.ObjectId})\n")
    return result.ObjectId

# def SSSetFirst(ss1:SelectionSet):
#     ids = ss1.GetObjectIds()
#     ed.SetImpliedSelection(ids)

# def SSIdsFirst(ids):
#     ed.SetImpliedSelection(ids)

def SSGet(dxfcode_filter_list=[]): # [[0, "Circle"], [0, "Line"]]
    # value = [TypedValue(System.Int32(0), "Circle")] # == AutoLisp (DxfCode . "Circle") 
    typevalue_list = []
    for [dxfcode, checkchar] in dxfcode_filter_list:
        typevalue_list.append(TypedValue(System.Int32(dxfcode), checkchar))
    # options = PromptSelectionOptions()
    # print(typevalue_list)
    filter = SelectionFilter(typevalue_list)
    result = ed.GetSelection(filter) # PromptSelectionResult
    ss1 = result.Value # SelectionSet
    # ids = ss1.GetObjectIds()
    # ss2 = SelectionSet.FromObjectIds([ids[-1]])
    Highlight(ss1)
    return ss1 


def Highlight(ss1:SelectionSet):
    if ss1 == None: return 
    dlock = doc.LockDocument()
    trans = db.TransactionManager.StartTransaction()
    for objid in ss1.GetObjectIds():
        objref = trans.GetObject(objid, OpenMode.ForRead)
        objref.Highlight()
    # trans.Commit()
    dlock.Dispose()


def GetEntityBound(objid:ObjectId):
    dlock = doc.LockDocument()
    trans = db.TransactionManager.StartTransaction() 
    objref = trans.GetObject(objid, OpenMode.ForRead)
    extend = objref.GeometricExtents
    # trans.Commit()
    dlock.Dispose()
    point1 = extend.MinPoint
    point2 = extend.MaxPoint
    return [point1.X, point1.Y, point1.Z], [point2.X, point2.Y, point2.Z]

def GetEntityBoundXY(objid:ObjectId):
    [x1,y1,z1],[x2,y2,z2] = GetEntityBound(objid)
    return [x1, y1, 0], [x2, y2, 0]

def GetEntityBoundCenterXY(objid:ObjectId):
    [x1,y1,z1],[x2,y2,z2] = GetEntityBoundXY(objid)
    return [(x1+x2)/2, (y1+y2)/2, 0]


def GetBound(ss1:SelectionSet):
    extend = Extents3d()
    dlock = doc.LockDocument()
    trans = db.TransactionManager.StartTransaction() 
    for objid in ss1.GetObjectIds():
        objref = trans.GetObject(objid, OpenMode.ForRead)
        extend.AddExtents(objref.GeometricExtents) 
    # trans.Commit()
    dlock.Dispose()
    point1 = extend.MinPoint
    point2 = extend.MaxPoint
    return [point1.X, point1.Y, point1.Z], [point2.X, point2.Y, point2.Z]


def GetBoundXY(ss1:SelectionSet):
    [x1,y1,z1],[x2,y2,z2] = GetBound(ss1)
    return [x1, y1, 0], [x2, y2, 0]


def GetBoundCenterXY(ss1:SelectionSet):
    [x1,y1,z1],[x2,y2,z2] = GetBoundXY(ss1)
    return [(x1+x2)/2, (y1+y2)/2, 0]



def GetLWPolyLinePointList(objid:ObjectId):
    dlock = doc.LockDocument()
    trans = db.TransactionManager.StartTransaction() # if works !!!
    pline = trans.GetObject(objid, OpenMode.ForRead)
    result = []
    for i in range(pline.NumberOfVertices):
        point = pline.GetPoint2dAt(i)
        result.append([point.X, point.Y, 0])
    flag = pline.Closed
    # print(result)
    dlock.Dispose()
    if flag: return result + result[0:1]
    return result


def GetString(string=""):
    if string == "": string = "请输入字符串: "
    result = ed.GetString(string) # PromptResult # (OK,4000x123) # [result.Status, result.StringResult]
    Prompt("\n")
    # print("GetString:", result)
    if result.StringResult == "": return None
    if result.Status == PromptStatus.OK: return result.StringResult
    return None


def GetInt(default_int:int, string=""):
    if string == "": string = "请输入整数: "
    options = PromptIntegerOptions(string)
    options.DefaultValue = default_int
    result = ed.GetInteger(options) # PromptResult # (OK,4000x123) # [result.Status, result.StringResult]
    Prompt("\n")
    # print("GetInt:", result)
    if result.Status == PromptStatus.OK: return result.Value
    return None




# GetPoint: ((OK,),(225.037456173945,41.6177530027926,0))
# GetPoint: ((Cancel,),(0,0,0))
def GetPoint(string="", pt0=[]):
    if string == "": string = "请选择顶点: "
    options = PromptPointOptions(string)
    if pt0 != []: options.BasePoint = ToPoint3d(pt0)
    result = ed.GetPoint(options)
    Prompt("\n")
    # print("GetPoint:", result)
    if result.Status == PromptStatus.OK: return [result.Value.X, result.Value.Y, result.Value.Z]
    return None

def GetCorner(string, base_point):
    if string == "": string = "请选择顶点: "
    result = ed.GetCorner(string, ToPoint3d(base_point))
    return [result.Value.X, result.Value.Y, result.Value.Z]

# GetDouble: ((OK,),45) # 回车 or 右键
# GetDouble: ((Cancel,),0) # ESC
def GetDouble(default_double:float, string=""):
    if string == "": string = "请输入数值: "
    options = PromptDoubleOptions(string)
    options.DefaultValue = default_double
    result = ed.GetDouble(options) 
    Prompt("\n")
    # print("GetDouble:", result)
    if result.Status == PromptStatus.OK: return result.Value
    return None


def GetDoubleListLimitCount(count=10):
    列表 = []
    for i in range(count):
        result = ed.GetDouble(f"请输入第{i+1}个数据:") 
        if result.Value == 0: break
        if result.Status == PromptStatus.OK: 列表.append(result.Value)
        else: break
    if 列表 == []: return None
    return 列表

 

# (command "style" "新字体样式名称" "字体文件名称" 0 1 0 0 0 0)
# (command "selectall")
# (command "chprop" "style" "新字体样式名称" "")
def AddFontStyle(style_name:str, font_name:str):
    Command(["-style", style_name, font_name, System.Int32(0), System.Int32(1), System.Int32(0), "N", "N"])


def ChangeFontStyle(style_name:str, new_font_name:str):
    AddFontStyle(style_name, new_font_name)



# (command "-style" "mystyle" "txt.shx,gbcbig.shx" 8 1 0 "N" "N" "N" )
# (command "-style" "mystyle" "Times New Roman" 8 1 0 "N" "N")
# 这两个都能正确运行，但很明显，第一行最后是三个N，第二行才两个，但不知道为什么这样，可能是“语境”吧
def ChangeStandardFontStyle(new_font_name:str):
    Command(["-style", "Standard", new_font_name, System.Int32(0), System.Int32(1), System.Int32(0), "N", "N"])


# GetSelection() 用户在图形中选择实体
# SelectAll()   选择所有实体
# SelectCrossingWindow() 选择窗口及和窗口四边相交的实体
# SelectCrossingPolygon 选择多边形中及和多边形相交的实体
# SelectFence 栏选
# SelectImplied 选择当前图形中已经选择的实体
# SelectLast 选择图形中最后一盒绘制的实体
# SelectPrevious 选择上一个选择集
# SelectWindows 选择窗口中的实体
# SelectWindowsPolygon 选择多边形中的实体
# SelectCrossingWindow 通过点坐标选择图形


# ObjectId plobj = Autodesk.AutoCAD.Internal.Utils.None
# 值	说明
# -5	APP：永久反应器链
# -4	APP：条件运算符（仅与 ssget 一起使用）
# -3	APP：扩展数据 (XDATA) 标记（固定）
# -2	APP：图元名参照（固定）
# -1	APP：图元名。每次打开图形时，图元名都会发生变化，从不保存（固定）
# 0	表示图元类型的字符串（固定）
# 1	图元的主文字值
# 2	名称（属性标记、块名等）
# 3-4	其他文字或名称值
# 5	图元句柄；最多 16 个十六进制数字的字符串（固定）
# 6	线型名（固定）
# 7	文字样式名（固定）
# 8	图层名（固定）
# 9	DXF：变量名称标识符（仅在 DXF 文件的 HEADER 段中使用）
# 10	主要点；直线或文字图元的起点、圆的圆心，等等。DXF：主要点的 X 值（后跟 Y 和 Z 值代码 20 和 30）。APP：三维点（三个实数的列表）
# 11-18	其他点。DXF：其他点的 X 值（后跟 Y 值代码 21-28 和 Z 值代码 31-38）。APP：三维点（三个实数的列表）
# 20, 30	DXF：主要点的 Y 值和 Z 值
# 21-28,31-37	DXF：其他点的 Y 值和 Z 值
# 38	DXF：如果非零，则为图元的标高
# 39	如果非零，则为图元的厚度（固定)
# 40-48	双精度浮点值（文字高度、缩放比例等）
# 48	线型比例；双精度浮点标量值；默认值适用于所有图元类型
# 49	重复的双精度浮点值。一个图元的可变长度表（例如，LTYPE 表中的虚线长度）中可能会出现多个 49 组。7x 组始终出现在第一个 49 组之前，用以指定表的长度
# 50-58	角度（在 DXF 文件中以度为单位，在 AutoLISP 和 ObjectARX 应用程序中以弧度为单位）
# 60	图元可见性；整数值；未赋值或值为 0 时表示可见；值为 1 时表示不可见
# 62	颜色号（固定）
# 66	“图元跟随”标志（固定）
# 67	空间 — 模型空间或图纸空间（固定）
# 68	APP：指示视口是处于打开状态但在屏幕上完全不可见，还是未激活或处于关闭状态
# 69	APP：视口标识号
# 70-78	整数值，例如重复计数、标志位或模式
# 90-99	32 位整数值
# 100	子类数据标记（将派生类名作为字符串）从其他具体类派生的所有对象和图元类必须具有此标记。子类数据标记用于分离由同一对象的继承链中的不同类定义的数据。对于从 ObjectARX 派生的每个不同的具体类的 DXF 名称来说，这是必须满足的额外要求（参见子类标记）
# 102	控制字符串，后跟“{<任意名称>”或“}”。与扩展数据 1002 组码类似，不同之处在于当字符串以“{”开始时，其后可跟任意字符串，字符串的解释取决于应用程序。唯一允许的另外一个控制字符串是作为组结束符的“}”。除了执行图形核查操作期间外，AutoCAD 不会解释这些字符串。它们供应用程序使用
# 105	DIMVAR 符号表条目的对象句柄
# 110	UCS 原点（仅当将代码 72 设置为 1 时才显示）DXF：X 值；APP：三维点
# 111	UCS X 轴（仅当将代码 72 设置为 1 时才显示）DXF：X 值；APP：三维矢量
# 112	UCS Y 轴（仅当将代码 72 设置为 1 时才显示）DXF：X 值；APP：三维矢量
# 120-122	DXF：UCS 原点的 Y 值，UCS X 轴和 UCS Y 轴
# 130-132	DXF：UCS 原点的 Z 值，UCS X 轴和 UCS Y 轴
# 140-149	双精度浮点值（例如点、标高和 DIMSTYLE 设置）
# 170-179	16 位整数值，例如表示 DIMSTYLE 设置的标志位
# 210	拉伸方向（固定）DXF：拉伸方向的 X 值APP：三维拉伸方向矢量
# 220, 230	DXF：拉伸方向的 Y 值和 Z 值
# 270-279	16 位整数值
# 280-289	16 位整数值
# 290-299	布尔标志值
# 300-309	任意字符串
# 310-319	具有相同表示和 1004 组码限制的任意二进制块：用最大长度为 254 个字符的十六进制字符串表示最大长度为 127 个字节的数据块
# 320-329	任意对象句柄；“按原样”获取的句柄值。它们在 INSERT 和 XREF 操作期间不进行转换
# 330-339	软指针句柄；指向同一个 DXF 文件或图形中的其他对象的任意软指针。在 INSERT 和 XREF 操作期间进行转换
# 340-349	硬指针句柄；指向同一个 DXF 文件或图形中的其他对象的任意硬指针。在 INSERT 和 XREF 操作期间进行转换
# 350-359	软所有者句柄；指向同一个 DXF 文件或图形中的其他对象的任意软所有者指针。在 INSERT 和 XREF 操作期间进行转换
# 360-369	硬所有者句柄；指向同一个 DXF 文件或图形中的其他对象的任意硬所有者指针。在 INSERT 和 XREF 操作期间进行转换
# 370-379	线宽枚举值 (AcDb::LineWeight)。作为 16 位整数存储和移动。自定义非图元对象可以使用整个范围内的组码，但图元类只能在其表示中使用 371-379 DXF 组码，因为 AutoCAD 和 AutoLISP 都始终假定 370 组码是图元的线宽。这使 370 组码与其他“通用”图元字段具有相同的行为
# 380-389	PlotStyleName 类型枚举 (AcDb::PlotStyleNameType)。作为 16 位整数存储和移动。自定义非图元对象可以使用整个范围内的组码，但图元类只能在其表示中使用 381-389 DXF 组码，原因与上述线宽范围相同表示 PlotStyleName 对象的句柄值的字符串，本质上是硬指针，但范围不同，更容易处理向后兼容。作为对象 ID（在 DXF 文件中为句柄）和 AutoLISP 中的特殊类型存储和移动。自定义非图元对象可以使用整个范围内的组码，但图元类只能在其表示中使用 391-399 DXF 组码，原因与上述线宽范围相同
# 400-409	16 位整数
# 410-419	字符串
# 420-427	32 位整数值。与真彩色一同使用时，表示 24 位颜色值的 32 位整数。高阶字节（8 位）为 0；低阶字节为包含“蓝色”值 (0-255)、然后是“绿色”值的无符号字符；次高阶字节是“红色”值。将此整数值转换为十六进制值将得到以下位掩码：0x00RRGGBB。例如，红色==200、绿色==100 和蓝色==50 的真彩色为 0x00C86432，而在 DXF 中以十进制表示则为 13132850
# 430-437	字符串；用于真彩色时，则为表示颜色名称的字符串
# 440-447	32 位整数值。用于真彩色时，表示透明度值
# 450-459	长整数
# 460-469	双精度浮点值
# 470-479	字符串
# 999	DXF：999 组码指示后面的行是注释字符串。SAVEAS 不会在 DXF 输出文件中包含这样的组，但 OPEN 则包括这些组并忽略注释。可以使用 999 组在您编辑的 DXF 文件中包含注释
# 1000	扩展数据中的 ASCII 字符串（最多可以包含 255 个字节）
# 1001	扩展数据的注册应用程序名（最多可以包含 31 个字节的 ASCII 字符串）
# 1002	扩展数据控制字符串（“{”或“}”）
# 1003	扩展数据图层名
# 1004	扩展数据中的字节数据块（最多可以包含 127 个字节）
# 1005	扩展数据中的图元句柄；最多可以包含 16 个十六进制数字的字符串
# 1010	扩展数据中的点 DXF：X 值（后跟 1020 和 1030 组）APP：三维点
# 1020,1030	DXF：点的 Y 值和 Z 值
# 1011	扩展数据中的三维世界空间位置 DXF：X 值（后跟 1021 和 1031 组）APP：三维点
# 1021,1031	DXF：世界空间位置的 Y 值和 Z 值
# 1012	扩展数据中的三维世界空间位移DXF：X 值（后跟 1022 和 1032 组）APP：三维矢量
# 1022,1032	DXF：世界空间位移的 Y 值和 Z 值
# 1013	扩展数据中的三维空间方向DXF：X 值（后跟 1022 和 1032 组）APP：三维矢量
# 1023,1033	DXF：世界空间方向的 Y 和 Z 值
# 1040	扩展数据双精度浮点值
# 1041	扩展数据距离值
# 1042	扩展数据缩放比例
# 1070	扩展数据 16 位有符号整数
# 1071	扩展数据 32 位有符号长整数



    # if save: return
    # filepath = "F:\\CADdll\\Drawing1.dwg"
    # doc.Database.SaveAs(filepath, True, DwgVersion.Current, doc.Database.SecurityParameters)
    # doc = Application.DocumentManager.Open(filepath, False)
    # Application.DocumentManager.MdiActiveDocument = doc







# """
# def AutoCad():
#     global acad, space 
#     acad = win32com.client.Dispatch("Autocad.Application")
#     acad.Visible = True 
#     space = Space()
#     pass



















# """
# CAD entity Object drawing
# """

# def AddPoint(x,y,z=0):
#     pass
#     return point 



# def AddLwpline(ptlist):
#     """
#     LightWeight Poly line, this method requires the group of 2D vertex coordinates(that is x, y), 
#     This method is recommended to draw line
#     """
#     vertexCoord = []
#     for pt1 in ptlist:
#         x, y = pt1[0:2]
#         vertexCoord.append(x)
#         vertexCoord.append(y)
#     lwpline = space.AddLightWeightPolyline(VtVertex(*vertexCoord))
#     return lwpline

# def AddLwpline3D(ptlist):
#     """
#     LightWeight Poly line, this method requires the group of 2D vertex coordinates(that is x, y), 
#     This method is recommended to draw line
#     """
#     vertexCoord = []
#     for x, y, z in ptlist:
#         vertexCoord.append(x)
#         vertexCoord.append(y)
#         vertexCoord.append(z)
#     lwpline = space.AddLightWeightPolyline3D(VtVertex(*vertexCoord))
#     return lwpline


# def AddRect(ptlst): # ptlst = [pt1, po1, po2, pt2]
#     rect1 = AddLwpline(ptlst)
#     rect1.Closed = True
#     return rect1


# def AddCircle(centerPnt, radius):
#     """
#     add a circle, centerPnt's type is Apoint
#     """
#     circle = space.AddCircle(Apoint(*centerPnt), radius)
#     return circle

# def AddCircleDiameter(centerPnt, diameter):
#     """
#     add a circle, centerPnt's type is Apoint
#     """
#     radius = diameter / 2
#     circle = space.AddCircle(Apoint(*centerPnt), radius)
#     return circle



# def AddArc(centerPnt, radius, startrad, finalrad):
#     """
#     add an arc, startAngle and endAngle are both in the form of degree
#     """
#     arc = space.AddArc(Apoint(*centerPnt), radius, startrad, finalrad)
#     return arc 

# def AddArcCenterAngle(centerPnt, radius, startAngle, finalAngle):
#     """
#     add an arc, startAngle and endAngle are both in the form of degree
#     """
#     arc = space.AddArc(Apoint(*centerPnt), radius, AngletoRad(startAngle), AngletoRad(finalAngle))
#     return arc 




# def Point3toRad(center, pt2, pt3):
#     # 算法错误，不可用
#     x1, y1 = center[0], center[1]
#     x2, y2 = pt2[0], pt2[1]
#     x3, y3 = pt3[0], pt3[1]

#     ab = math.sqrt((x2 - x1) ** 2 + (y2 - y1) ** 2)
#     ac = math.sqrt((x3 - x1) ** 2 + (y3 - y1) ** 2)
#     dot_product = (x2 - x1) * (x3 - x1) + (y2 - y1) * (y3 - y1)
#     rad = math.acos(dot_product / (ab * ac))
#     if y3 < y2: rad = -rad
#     return rad # return math.degrees(angle)



# def AddArc3point(pt1, pt2, pt3):
#     x1, y1 = pt1[0], pt1[1]
#     x2, y2 = pt2[0], pt2[1]
#     x3, y3 = pt3[0], pt3[1]

#     A=x1*(y2-y3)-y1*(x2-x3)+x2*y3-x3*y2
#     B=(x1*x1+y1*y1)*(y3-y2)+(x2*x2+y2*y2)*(y1-y3)+(x3*x3+y3*y3)*(y2-y1)
#     C=(x1*x1+y1*y1)*(x2-x3)+(x2*x2+y2*y2)*(x3-x1)+(x3*x3+y3*y3)*(x1-x2)
#     D=(x1*x1+y1*y1)*(x3*y2-x2*y3)+(x2*x2+y2*y2)*(x1*y3-x3*y1)+(x3*x3+y3*y3)*(x2*y1-x1*y2)
#     x=-B/(2*A)
#     y=-C/(2*A)
#     r=math.sqrt((B*B+C*C-4*A*D)/(4*A*A))
#     center = [x, y]
#     pt0 = [x+r, y]

#     startrad = Point3toRad(center, pt0, pt1)
#     finalrad = Point3toRad(center, pt0, pt3)
#     sideflag = WhichSideOfLine(pt1, pt3, pt2)
#     if sideflag > 0: # 在左侧
#         AddArc(center, r, finalrad, startrad)
#     else:
#         AddArc(center, r, startrad, finalrad)


# # def CommandArc3point(pt1, pt2, pt3): # Error   
# #     SendCommand(f"arc {pt1[0]} {pt1[1]} {pt1[2]} {pt2[0]} {pt2[1]} {pt2[2]} {pt3[0]} {pt3[1]} {pt3[2]} ")


# def AddTable(InsertionPoint, NumRows, NumColumns, RowHeight, ColWidth):
#     """
#     add a table in ModelSpace.
#     InsertionPoint:Apoint type.
#     NumRows:The number of rows in the table. 
#     NumColumns:The number of columns in the table.
#     RowHeight :The height of the rows in the table. 
#     ColWidth :The width of the columns in the table. 
#     """
#     return space.AddTable(Apoint(*InsertionPoint), NumRows, NumColumns, RowHeight, ColWidth)

# def AddSpline(fitpoints, startTanPt = None, endTanPt = None):
#     """
#     fitpoints are array of 3D coordinates of points, such as (1, 2, 3, 4, 5, 6)
#     startTan is the starting vector which is the type of Apoint, and the same is endTan.
#     """
#     spline = space.addSpline(VtVertex(*fitpoints), Apoint(*startTanPt), Apoint(*endTanPt))
#     return spline

# def AddEllipse(centerPnt, majorAxisPt, radiusRatio):
#     """
#     add ellipse, the type of majorAxis is Apoint
#     """
#     ellipse = space.addEllipse(Apoint(*centerPnt), Apoint(*majorAxisPt), radiusRatio)
#     return ellipse

# def AddHatch(patternType, patterName, associative, outLoopTuple, innerLoopTuple = None):
#     """
#     The note of arguments can be seen as below:
#     (1)patternType is the built-in integer constants which can be got by win32com.client.constants.X, here x can be 
#     acHatchPatternTypeDefined(it means that from standard drawing file Acad.Pat to hatch, and the integer is 1), 
#     acHatchPatternTypeUserDefined(it means that from the current linetype to hatch, and the integer is 0), 
#     acHatchPatternTypeCustomDefined(it means that from user-defined drawing file .Pat to hatch, and the integer is 2)
#     (2)patterName is a string specifying the hatch pattern name, such as "SOLID", "ANSI31"
#     (3)associative is boolean. If it is True, when the border is modified, the hatch pattern will adjust automatically 
#     to keep in the modified border.
#     (4)outLoop is a sequence of object, such as line, circle, etc. For example, outLoopTuple = (circle1, ), or outLoopTuple = (line1, line2, 
#     line3).
#     (5)innerLoop is the same with outLoop
#     """
#     hatch = space.AddHatch(patternType, patterName, associative)
#     out = VtObject(*outLoopTuple)
#     hatch.AppendOuterLoop(out)
#     if innerLoopTuple:
#         inn = VtObject(*innerLoopTuple)
#         hatch.AppendInnerLoop(inn)
#     hatch.Evaluate()
#     return hatch 

# def AddSolid(pnt1, pnt2, pnt3, pnt4):
#     """
#     Creates a 2D solid polygon. 
#     pnt1, pnt2, pnt3, pnt4:Apoint type
#     """
#     return space.AddSolid(Apoint(*pnt1), Apoint(*pnt2), Apoint(*pnt3), Apoint(*pnt4))

# def AboutEntityObject():

#     """
#     <This method is created only for the noting purpose>
#     About editting autocad entity object:Users shall consult the acadauto.chm located in  C:\\Program Files
#     \\Common Files\\Autodesk Shared\\acadauto.chm for exact supported property in terms of every kind of cad 
#     entity.Some commen property and method has been summed up as below:
#         (1)Commen Property:
#             (a)object.color = X
#             X:built-in contant, such as win32com.client.constants.acRed.Here, color is lowercase.
#             (b)object.Layer = X
#             X:string, the name of the layer
#             (c)object.Linetype = X
#             X:string, the name of the loaded linetype
#             (d)object.LinetypeScale = X
#             X:float, the linetype scale
#             (e)object.Visible = X
#             X:boolean, Determining whether the object is visible or invisible
#             (f)object.EntityType
#             read-only, returns an integer.
#             (g)object.EntytyName
#             read-only, returns a string
#             (h)object.Handle
#             read-only, returns a string
#             (i)object.ObjectID
#             read-only, returns a long integer
#             (j)object.Lineweight = X
#             X:built-in constants, (For example, win32com.client.constants.acLnWt030(0.3mm), acLnWt120 is
#             1.2mm, and the scope of lineweight is 0~2.11mm), or acByLayer(the same with the layer
#             where it lies), acByBlock, acBylwDefault.

#         (2)Commen Method:
#             (a)Copy
#             RetVal = object.Copy
#             RetVal: New created object
#             object:Drawing entity, such as Arc, Line, LightweithPolyline, Spline, etc.

#             (b)Offset
#             RetVal = object.Offset(Distance)
#             RetVal:New created object tuple
#             Distance:Double, positive or negative
#             object:Drawing entity, such as Arc, Line, LightweithPolyline, Spline, etc.

#             (c)Mirror
#             RetVal = object.Mirror(point1, point2)
#             RetVal:mirror object
#             point1, point2:end of mirror axis, Apoint type.
#             object:Drawing entity, such as Arc, Line, LightweithPolyline, Spline, etc.

#             (d)ArrayPolar
#             RetVal = object.ArrayPolar(NumberOfObject, AngleToFill, CenterPoint)
#             RetVal:New created object tuple
#             NumberOfObject:integer, the number of array object(including object itself)
#             AngleToFill:Double, rad angle, positive->anticlockwise, negative->clockwise
#             CenterPoint:Double, Apoint type. The center of the array.
#             object:Drawing entity, such as Arc, Line, LightweithPolyline, Spline, etc.

#             (e)ArrayRectangular
#             RetVal = object.ArrayRectangular(NumberOfRows, NumberOfColumns, NumberOfLevels, 
#             DistBetweenRows, DistBetweenColumns, DistBetweenLevels)
#             RetVal:new created object tuple
#             NumberOfRows, NumberOfColumns, NumberOfLevels:integer, the number of row, column, level, 
#             if it is the plain array that is performed, NumberOfLevels = 1
#             DistBetweenRows, DistBetweenColumns, DistBetweenLevels:Double, the distance between rows, 
#             columns, levels respectively.When NumberOfLevels = 1, DistBetweenLevels is valid but still
#             need to be passed
#             object:Drawing entity, such as Arc, Line, LightweithPolyline, Spline, etc.

#             (f)Move
#             object.Move(point1, point2)
#             object:Drawing entity, such as Arc, Line, LightweithPolyline, Spline, etc.
#             point1, point2:Double, Apoint type. The moving vector shall be determined by the
#             two points and point1 is the start point, point2 is the end point.

#             (g)Rotate
#             object.Rotate(BasePoint, RotationAngle)
#             object:Drawing entity, such as Arc, Line, LightweithPolyline, Spline, etc.
#             BasePoint:Double, Apoint type.The rotation basepoint.
#             RotationAngle:Double, rad angle.

#             (h)ScaleEntity
#             object.ScaleEntity(BasePoint, ScaleFactor)
#             object:Drawing entity, such as Arc, Line, LightweithPolyline, Spline, etc.
#             BasePoint:Double, Apoint type.The scale basepoint.
#             ScaleFactor:Double, Apoint type.

#             (i)Erase
#             object.Erase()
#             object:Choosed set
#             Delete all entity in the choosen scope

#             (J)Delete
#             object.Delete()
#             object:specified entity, as for set object, such as modelSpace set and layerSet , this
#             method is valid.

#             (k)Update
#             object.Update()
#             update object after some kind of the objects' editing.

#             (L)color
#             object.color
#             Here attention please, it is color, Not Color.(lowercase)

#             (M)TransformBy
#             object.TransformBy(transformationMatrix)
#             object:Drawing entity, such as Arc, Line, LightweithPolyline, Spline, etc.
#             transformationMatrix:4*4 Double array, need to be passed to ArrayTransform() method to be the required type 


#         (3)Refer to Object

#             (a)HandleToObject
#             RetVal = object.HandleToObject(Handle)
#             Retval:the entity object corresponding to Handle
#             object:Document object
#             Handle: the handle of entity object

#             (b)ObjectIdToObject
#             RetVal = object.ObjectIdToObject(ID)
#             RetVal:the entity object corresponding to ID
#             object:Document object
#             ID: the identifier of object 
#     """
#     pass










# """
# AutoCad Data Type
# """

# def Apoint(x, y, z = 0):
#     """
#     Converts x, y, z into required float array as the arguments of coordinates of a point
#     """
#     return win32com.client.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_R8, (x, y, z))  # | means the type combination

# def ArrayTransform(x):
#     """
#     x: any kind of array in python, such as ((1, 2, 3), (1, 2, 3), (1, 2, 3))
#     """
#     return win32com.client.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_R8, x) 

# def VtVertex(*args):
#     """
#     Converts 2D coordinates of a serial points into the required float array
#     """
#     return win32com.client.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_R8, args)

# def VtObject(*obj):
#     """
#     converts obj in python into required obj array
#     """
#     return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, obj)

# def VtFloat(list):
#     """converts list in python into required float"""
#     return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, list)

# def VtInt(list):
#     """converts list in python into required int"""
#     return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, list)

# def VtVariant(list):
#     """converts list in python into required variant"""
#     return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, list)

# def AngletoRad(angle): # == math.radians(angle)
#     """
#     convert degree to rad
#     """
#     rad = angle*math.pi/180
#     return rad 

# def RadtoAngle(rad): # math.degrees(rad)
#     """
#     convert rad to degree
#     """
#     angle = 180*rad/math.pi 
#     return angle

# def FilterType(ftype):
#     """
#     ftype shall be tuple.Refer to DXF in acad_aag.chm to learn about DXF group code.
#     """
#     return win32com.client.VARIANT(pythoncom.VT_I2|pythoncom.VT_ARRAY, ftype)

# def FilterData(fdata):
#     """
#     fdata shall be tuple.Refer to DXF in acad_aag.chm to learn about DXF group code.
#     """
#     return win32com.client.VARIANT(pythoncom.VT_VARIANT|pythoncom.VT_ARRAY, fdata)

# class PycomError(Exception):
#     def __init__(info):
#         print(info)



# """
# version
# """
 
# def Version():
#     version_dict = {
#         '15.0':'AutoCAD2002', 
#         '16.0':'AutoCAD2004', 
#         '16.1':'AutoCAD2005', 
#         '16.2':'AutoCAD2006', 
#         '17.0':'AutoCAD2007', 
#         '17.1':'AutoCAD2008', 
#         '17.2':'AutoCAD2009', 
#         '18.0':'AutoCAD2010', 
#         '18.1':'AutoCAD2011', 
#         '18.2':'AutoCAD2012', 
#         '19.0':'AutoCAD2013', 
#         '19.1':'AutoCAD2014', 
#         '20.0':'AutoCAD2015', 
#         '20.1':'AutoCAD2016', 
#         '21.0':'AutoCAD2017', 
#         '22.0':'AutoCAD2018', 
#         '23.0':'AutoCAD2019', 
#         '23.1':'AutoCAD2020'
#     }
#     v = acad.version.split('s')[0]
#     return version_dict.get(v, 'UnKnown')

# """
# Application
# """
 
# def Space():
#     """
#     Automatically deciding the active layout is ModelSpace or PaperSpace.
#     """
#     if acad.ActiveDocument.ActiveLayout.ModelType:
#         return acad.ActiveDocument.ModelSpace
#     else:
#         return acad.ActiveDocument.PaperSpace

# def IsEarlyBind():
#     """
#     Test whether it is
#     :return:
#     """
#     if 'IAcadApplication' in str(type(acad)):
#         return True
#     else:
#         return False
    
# def TurnOnEarlyBind():
#     import os, sys
#     makepyPath = r'Lib\site-packages\win32com\client\makepy.py'
#     ExePath = os.path.split(sys.executable)[0]
#     MakePyPath = os.path.join(ExePath, makepyPath)
#     os.execl(sys.executable, 'python', MakePyPath)

# def AppPath():
#     """
#     return Autocad Application path
#     """
#     return acad.Path

# def SendCommand(command):
#     """
#     Sends a command string from script
#     for example, to draw a circle whose center is 0, 0, 0, and radius is 100
#     >>>acad.SendCommand('circle 0, 0, 0 100 ') #Please notice that there is a blank following 100, meaning the end of input
#     """
#     acad.ActiveDocument.SendCommand(command)





# def QdimEntlast(pt1):
#     x, y = pt1[0:2]
#     command = f"qdim (entlast) \"\" {x} {y} 0"
#     SendCommand(command)



# """
# Registered Applications
# """
 
# def RApps():
#     """
#     return registered applications collections, it has `Add(str)` and `Item(int)` method, `Count` property
#     """
#     return acad.ActiveDocument.RegisteredApplications

 
# def RAppNames():
#     """
#     return registerd application names list
#     """
#     rApps = RApps()
#     names = []
#     for item in range(rApps.Count):
#         names.append(rApps.Item(item).Name)
#     return names

# def SetXData(entity, xdataPairs):
#     """
#     set XData for entity
#     xdataPairs:list containing tuple which represent xdataType and xdata.
#     >>>circle = acad.AddCircle(Apoint(0, 0), 20)
#     >>>acad.SetXdata(circle, [(1001, 'test'), (1000, 'this is an example')])
#     """
#     xdataType = []
#     xdataValue = []
#     for i, j in xdataPairs:
#         xdataType.append(i)
#         xdataValue.append(j)
#     entity.SetXData(FilterType(xdataType), FilterData(xdataValue))
    

# """
# Layouts
# """
 
# def LayoutNames():
#     layoutnames = {}
#     for i in range(acad.ActiveDocument.Layouts.Count):
#         layoutnames[acad.ActiveDocument.Layouts.Item(i).Name] = i
#     return layoutnames

# def EnterLayout(name_or_index):
#     """
#     enter layout, so the active layout may become model, or one paperspace.
#     name_or_index:string, the name of layout;int, the index of layout
#     For example:
#     >>>acad.EnterLayout('model') #Enter the modelSpace
#     >>>acad.EnterLayout('onePaperSpace') #Enter the paperSpace named onePaperSpace
#     >>>acad.EnterLayout('notExist') # Will raise an error
#     >>>acad.EnterLayout(0) # Enter the modelSpace
#     """
#     if isinstance(name_or_index, int):
#         acad.ActiveDocument.ActiveLayout = acad.ActiveDocument.Layouts.Item(name_or_index)
#     if isinstance(name_or_index, str):
#         if name_or_index in ['Model', '模型']:
#             acad.ActiveDocument.ActiveSpace = win32com.client.constants.acModelSpace
#             return None 
#         index = LayoutNames.get(name_or_index, -1)
#         acad.ActiveDocument.ActiveLayout = acad.ActiveDocument.Layouts.Item(index)
#         return None 

# """
# System variable
# """
# def SetVariable(name, value):
#     """
#     name:string
#     """
#     acad.ActiveDocument.SetVariable(name, value)
# def GetVariable(name):
#     return acad.ActiveDocument.GetVariable(name)


# """
# File processing
# """

# def OpenFile(path):
#     """
#     open a dwg file in the path
#     """
#     acad.Documents.Open(path)

# def CreateNewFile():
#     """
#     Create a new dwg file, by default, the name is Drawing1.dwg
#     """
#     acad.Documents.Add()

# def SaveFile():
#     """
#     save file
#     """
#     acad.ActiveDocument.Save()

# def SaveAsFile(path):
#     """
#     save as file
#     """
#     acad.ActiveDocument.SaveAs(path)

# def Close():
#     """
#     Close current file
#     :return:
#     """
#     acad.ActiveDocument.Close()

# def PurgeAll():
#     """
#     Removes unused named references such as unused blocks or layers from the document.
#     This method is the equivalent of entering purge at the Command prompt, selecting the 
#     All option, and then choosing Yes to the "Purge Everything?" prompt. 
#     """
#     acad.ActiveDocument.PurgeAll()

# def Regen(enum):
#     """
#     Regenerates the entire drawing and recomputes the screen coordinates and view resolution for all objects.
#     enum:
#         0:Regenerates only the active viewport
#         1:Regenerates all viewports on the document
#     """
#     acad.ActiveDocument.Regen(enum)


# def OpenedFilenames():
#     """
#     :return: list, all opened filenames
#     """
#     names = []
#     for i in range(OpenedFilenumbers):
#         names.append(acad.Documents.Item(i).Name)
#     return names

# def OpenedFilenumbers():
#     """
#     :return: the number of opened file
#     """
#     return acad.Documents.Count

# def GetOpenedFile(file):
#     """
#     Return already opened file whose index is index or name is name as
#     the Current file
#     :param file: int, or string, the number of index or the name of file targeted to be set as the current file
#     """
#     if isinstance(file, str):
#         index = OpenedFilenames.index(file)
#     elif isinstance(file, int):
#         index = file
#     else:
#         raise PycomError('Type of file in GetOpenedFile is wrong ')
#     return acad.Documents.Item(index)

# def ActivateFile(file):
#     """
#     Activate already opened file whose index is index or name is name as
#     the Current file
#     :param file: int, or str, the number of index or the name of file targeted to be set as the current file
#     """
#     if isinstance(file, str):
#         index = OpenedFilenames.index(file)
#     elif isinstance(file, int):
#         index = file
#     else:
#         raise PycomError('Type of file in ActivateFile() is wrong')
#     acad.Documents.Item(index).Activate()

# def DeepClone(objects, Owner = None, IDPairs = win32com.client.VARIANT(pythoncom.VT_VARIANT, ())):
#     """
#     Deep clone objects from current file to specified file's ModelSpace
#     :param objects: objects needed to be deep cloned.Type: IAcadSelectionSet(selection sets), tuple of entity object
#     :param Owner:specified opened file name, Type: string;Or the index of specified opened file name.
#     :param ID:IDPairs.Default value has been set.
#     :return:tuple of deep cloned object
#     For example:
#     >>from pycomcad import *
#     >>acad = Autocad()
#     >>te1 = acad.AddCircle(Apoint(0, 0, 0), 200)
#     >>te2 = acad.AddCircle(Apoint(100, 100, 0), 200)
#     >>acad.CreateNewFile()
#     >>acad.ActivateFile(0)
#     >>result = acad.DeepClone((te1, ), 1) # Deep Clone one object, notice the naunce between (te1, )and (te1), the latter one is int.
#     >>result[0][0].Move(Apoint(0, 0, 0), Apoint(100, 100, 0))
#     >>acad.CurrentFilename
#     >>slt = acad.GetSelectionSets('slt1')
#     >>slt.SelectOnScreen()
#     >>result1 = acad.DeepClone(slt, 'Drawing2.dwg')
#     """
#     if isinstance(objects, tuple):
#         if not objects:
#             raise PycomError('Objects in DeepClone() is empty tuple ')
#         else:
#             obj = VtObject(*objects)
#     elif 'IAcadSelectionSet' in str(type(objects)):
#         if objects.Count == 0:
#             raise PycomError('SelectionSets in DeepClone() is empty')
#         else:
#             obj = []
#             for i in range(objects.Count):
#                 obj.append(objects.Item(i))
#             obj = VtObject(*obj)
#     else:
#         raise PycomError('Type of objects in DeepClone() is wrong')
#     if not Owner:
#         return acad.ActiveDocument.CopyObjects(obj)
#     else:
#         try:
#             newOwnerDoc = GetOpenedFile(Owner)   
#             if newOwnerDoc.ActiveLayout.ModelType:  # make deepclone method can be applied to paperspace or modelspace automatically
#                 newOwner = newOwnerDoc.ModelSpace
#             else:
#                 newOwner = newOwnerDoc.PaperSpace
#         except:
#             raise PycomError('File %s is not opened'% Owner)
#         return acad.ActiveDocument.CopyObjects(obj, newOwner, IDPairs)


# def CurrentFilename():
#     """
#     return: str, the name of current file name
#     """
#     return acad.ActiveDocument.Name


# def FilePath():
#     """
#     return current file path
#     """
#     return acad.ActiveDocument.Path

# def IsSaved():
#     """
#     Specifies if the document has any unsaved changes
#     return:True: The document has no unsaved changes.
#             False: The document has unsaved changes.
#     """
#     return acad.ActiveDocument.Saved



# """
# Zoom
# """

# def ZoomExtents():
#     acad.ZoomExtents()

# def ZoomAll():
#     acad.ZoomAll()

# """
# precise-drawing setting
# """

# def GridOn(boolean):
#     """
#     grid-on
#     """
#     acad.ActiveDocument.ActiveViewport.GridOn = boolean
#     acad.ActiveDocument.ActiveViewport = acad.ActiveDocument.ActiveViewport

# def SnapOn(boolean):
#     """
#     snap-on
#     """
#     acad.ActiveDocument.ActiveViewport.SnapOn = boolean
#     acad.ActiveDocument.ActiveViewport = acad.ActiveDocument.ActiveViewport
    



# """
# Refer and select entity
# """

# def Handle2Object(handle):
#     """
#     handle:entities or its reference objects' handle.
#     return the object of specified handle
#     """
#     return acad.ActiveDocument.HandleToObject(handle)


# def GetEntityByItem(i):
#     """
#     Refere to entity by its index location
#     """
#     return space.Item(i)

# def GetSelectionSets(setname):
#     """
#     setname:string, the name of selection set.
#     There are 2 steps to select entity object:
#     (1) create selection set 
#     (2)Add entity into set
#     Also note that: one set once has been created, 
#     it can never be created again, unless it is
#     deleted.
#     This method provides the first step.
#     For example:
#     >>>ft = [0, -4, 40, 8]  # define filter type
#     >>>fd = ['Circle', '> = ', 5, '0'] #define filter data
#     >>>ft = VtInt(ft) # data type convertion
#     >>>fd = VtVariant(fd) #data type convertion
#     >>>slt = acad.GetSelectionSets('slt') # Create selectionset object
#     >>>slt.SelectOnScreen(ft, fd) # select on screen
#     >>>slt.Erase() # Erase selected entity
#     >>>slt.Delete() # Delete selectionsets object
#     from select method:
#     >>>slt = acad.GetSelectionSets('slt1')
#     >>>slt.Select(Mode = win32com.client.constants.acSelectionSetAll, FilterType = ft, FilterData = fd) # Attention about the keyword arguments
#     (3) from SelectByPolygon method to automatically select entity
#     >>>pnt = acad.GetPoint()
#     >>>pnt1 = acad.GetPoint()
#     >>>pnt2 = acad.GetPoint()
#     >>>pnt3 = acad.GetPoint()  # select 4 points
#     >>>c = list(pnt)+list(pnt1)+list(pnt2)+list(pnt3)
#     >>>slt = acad.GetSelectionSets('test2')
#     >>>slt.SelectByPolygon(Mode = win32com.client.constants.acSelectionSetWindowPolygon, PointsList = VtVertex(*c))
#     """
#     return acad.ActiveDocument.SelectionSets.Add(setname)

# """
# There are 5 methods to add entity into selection set:

# (1)object.AddItems(Items)
# object:selection set
# Items:Variant tuple. For example, Items = VtObject((c1, c2)), where c1, c2 
# is the object being ready to join in selection set

# (2)object.Select(Mode[, Point1][, Point2][, FilterType][, FilterData])
# object:selection set
# Mode = win32com.client.constants.X
# X is as below:
# acSelectionSetWindow, acSelectionSetPrevious, acSelectionSetLast, 
# acSelectionSetAll
# Point1, Point2: 2 diagonal points defining a window
# FilterType, FilterData: DXF group code; filter type. 

# (3)object.SelectAtPoint(Point[, FilterType][, FilterData])
# object:selection set
# Point:Given point

# (4)object.SelectByPolygon(Mode, PointsLists[, FilterType][, FilterData])
# object:selection set
# Mode = win32com.client.constants.X
# X is as below:
# acSelectionSetFence, acSelectionSetWindowPolygon, acSelectionSetCrossingPolygon
# PointsLists:a serial points(3D) defining polygon
# FilterType, FilterData: DXF group code; filter type.

# (5)object.SelectOnScreen(filterType, filterData)
# object:selection set
# FilterType, FilterData: DXF group code; filter type.
# """

# """
# Filter Mechanism:
# DXF                  filter type
# 0              entity , such as Line, Circle, Ac, etc.
# 2              name of object (string)
# 5              entity handle
# 8              layer
# 60             visible of entity
# 62             color integer, 0->BYBLOCK, 256->BYLAYER, negative->closed layer
# 67             ignored or 0->ModelSpace, 1->PaperSpace

# DXF shall be passed into FilterType() in the form of tuple to be the required type, 
# while filter type shall be passed into FilterData() in the form of tuple to be the 
# required type.
# """


# """
# Deletion of selection set:
# (1)Clear:clear the selection set, the selection set still exists and the member entities still
# exist but they no longer belong to this selection set.

# (2)RemoveItems:the removed member entities still exist, but they no longer belong to this selection
# set.

# (3)Erase:delete all the member entities and the selection set itsel still exists.

# (4)Delete:delete the selection set itself, but the member entities still exist.

# """

# """
# Layer
# """
# def AddLayer(layername):
#     """
#     create new layer named layername(string)
#     """
#     return acad.ActiveDocument.Layers.Add(layername)


# def SetActivateLayer(layername):
#     """
#     Activate layer
#     layer:str or int, the index or the name of the being activated layer
#     """
#     acad.ActiveDocument.ActiveLayer = GetLayer(layername)


# def LayerNumbers():
#     """
#     return the number of layers in the active document
#     """
#     return acad.ActiveDocument.Layers.Count


# def LayerNames():
#     """
#     :return a list containing all layer names
#     """
#     a = []
#     for i in range(LayerNumbers):
#         a.append(acad.ActiveDocument.Layers.Item(i).Name)
#     return a
# def GetLayer(layer):
#     """
#     get an indexed layer
#     layer:int or string, the index or the name of layer which exists already.
#     """
#     if isinstance(layer, str):
#         index = LayerNames.index(layer)
#     elif isinstance(layer, int):
#         index = layer
#     else:
#         raise PycomError('Type of layer in GetLayer() is wrong')
#     return acad.ActiveDocument.Layers.Item(index)
 
# def Layers():
#     """
#     :return layer set object
#     """
#     return acad.ActiveDocument.Layers

# def ActiveLayer():
#     """
#     :return: ActiveLayer object
#     """
#     return acad.ActiveDocument.ActiveLayer
# """
# The state change and deletion of layer:
# (1)Obj.LayerOn = True/False
# Obj:Layer object
# closed or not, if it is  closed, new entity object can be created on layer, while it cannot be seen.
# (2)Obj.Freeze = True/False
# Obj:Layer object
# if freezed, the layer can neighter be shown or created entities on it.
# (3)Obj.Lock = True/False
# Obj:Layer object
# The entity on a locked layer can be shown, if the locked layer is activated, new entity can be 
# created there, but the entities cannot be edited or deleted.
# (4)Obj.Delete
# Obj:Layer object
# Delete any layer, except for cunnrent layer and 0 layer(default layer).

# The property of layer:
# (1)Obj.color = X
# X:built-in contant, such as win32com.client.constants.acRed
# (2)Obj.Linetype = X
# X:string, the name of loaded linetype
# (3)Obj.Name
# """
# """
# Linetype
# """
# def LoadLinetype(typename, filename = 'acad.lin'):
#     """
#     typename:string, the name of type needed to be load.such as 'dashed', 'center'
#     filename:string, the name of the file the linetype is in.'acad.lin', 'acadiso.lin'
#     """
#     acad.ActiveDocument.Linetypes.Load(typename, filename)

# def ActivateLinetype(typename):
#     """
#     typename:string, ensure the typename has been loaded successfully.
#     """
#     try:
#         acad.ActiveDocument.ActivateLinetype = acad.ActiveDocument.Linetypes.Item(typename)
#     except:
#         print('The typename has not been loaded')

# def ShowLineweight(TrueorFalse):
#     """
#     TrueorFalse:Boolean, determining whether the lineweight be shown or not
#     """
#     acad.ActiveDocument.Preferences.LineWeightDisplay = TrueorFalse

# def Linetypes():
#     """
#     return linetype set
#     """
#     return acad.ActiveDocument.Linetypes 
    


# """
# Block
# There are 3 steps as for the creation and reference about Block.
# (1)create a block, see  the following method CreateBlcok
# (2)The created blcok adds enity;
# Obj.AddX
# X can be entity object, text object, etc.
# (3)insert block, see the fowllowing method InsertBlcok

# Block Explode
# Obj.Explode()
# Obj:Reference Block object
# This method returns a tuple containing the exploded object

# Block attribute object
# Retval = blockObj.AddAttribute(height, mode, prompt, insertPoint, tag, value)
# blockObj:Block reference object
# Retval:Attribute object
# height:Double float, the height of text
# Mode:built-in constants, win32com.client.constants.X, and X is as the following
#     acAttributeModeInvisible:the attribute value is invisible
#     acAttributeModeConstant:constant attribute, cannot be editted
#     acAttributeModeVerify:when inserting block, prompt users to ensure the attribute value
#     acAttributeModePreset:when inserting block, use default attribute value, cannot be editted
#     These constants can be used as a combination

# GetAttribute method
#     To access an attribute reference of an inserted block, use the GetAttributes method. 
#     This method returns an array of all attribute references attached to the inserted block. 
# Retval = obj.GetAttributes()
# obj:Block reference object
# Retval:Block attribute object tuple
# Retval's 2 main property:(1)TagString(2)TextString
# Note:since Retval is a tuple, we may use len() method to get the number of the member in it
# """

# def CreateBlock(insertPnt, blockName):
#     """
#     insertPnt:Apoint type, the insertion base point
#     blockName:string, the name of the new-created block
#     """
#     return acad.ActiveDocument.Blocks.Add(Apoint(*insertPnt), blockName)

# def InsertRefBlock(insertPnt, blockName, Xscale = 1, Yscale = 1, Zscale = 1, Rotation = 0):
#     """
#     insertPnt:Apoint type, the insert point in the process of block insertion.
#     blockName:string, the inserted block name which has been created
#     """
#     return space.InsertBlock(Apoint(*insertPnt), blockName, Xscale, Yscale, Zscale, Rotation)

# """
# User-defined coordinate system
# Normally, users perform drawing work in WCS(world coordinate system).However, in some case, it is easy
# to draw in UCS(user coordinate system). In UCS, it's necessary to use coordinate transform, and the steps
# are as follow:
# (1)Create entity in WCS directly
# (2)Create UCS and get transform matrix of UCS by method GetUCSMatrix (here, also need array type conversion
#     by ArrayTransform method)
# (3)Transform the entity created in WCS to UCS through method TransformBy
# Also attention that after the transform perform , it's better to set the previous coordinate system.

# TransMatrix = ucsObj.GetUCSMatrix()
# TransMatrix:4*4 Double array, need to be passed to ArrayTransform() method to be the required type 
# ucsObj:UCS object

# TransformBy
# object.TransformBy(transformationMatrix)
# object:Drawing entity, such as Arc, Line, LightweithPolyline, Spline, etc.
# transformationMatrix:4*4 Double array, need to be passed to ArrayTransform() method to be the required type 
# """
# def CreateUCS(origin, xAxisPnt, yAxisPnt, csName):
#     """
#     origin:Apoint type, origin point of the new CS
#     xAxisPnt:Apoint type, one point directing the positive direction of x axis of the new CS
#     yAxisPnt:Apoint type, one point directing the positive direction of y axis of the new CS
#     csName:string, the name of the new CS
#     """
#     return acad.ActiveDocument.UserCoordinateSystems.Add(origin, xAxisPnt, yAxisPnt, csName)

# def ActivateUCS(ucsObj):
#     """
#     ucsObj: UCS object
#     """
#     acad.ActiveDocument.ActiveUCS = ucsObj

# def GetCurrentUCS():
#     """
#     Before activate the new created UCS, it's better to get the current UCS in case of activating
#     it after tasks in the activated new created UCS.
#     """
#     if acad.ActiveDocument.GetVariable('ucsname') == '':
#         origin = acad.ActiveDocument.GetVariable('ucsorg')
#         origin = ArrayTransform(origin)
#         xAxisPnt = acad.ActiveDocument.Utility.TranslateCoordinates(ArrayTransform(acad.ActiveDocument.GetVariable('ucsxdir')), 
#             win32com.client.constants.acUCS, win32com.client.constants.acWorld, 0)
#         xAxisPnt = ArrayTransform(xAxisPnt)
#         yAxisPnt = acad.ActiveDocument.Utility.TranslateCoordinates(ArrayTransform(acad.ActiveDocument.GetVariable('ucsydir')), 
#             win32com.client.constants.acUCS, win32com.client.constants.acWorld, 0)
#         yAxisPnt = ArrayTransform(yAxisPnt)
#         currCS = acad.ActiveDocument.UserCoordinateSystems.Add(origin, xAxisPnt, yAxisPnt, 'currentUCS')
#     else:
#         currCS = acad.ActiveDocument.ActiveUCS
#     return currCS

# def ShowUCSIcon(booleanOfUCSIcon, booleanOfUCSatOrigin):
#     """
#     show UCS Icon
#     booleanOfUCSIcon:boolean, Specifies if the UCS icon is on
#     booleanOfUCSatOrigin:boolean, Specifies if the UCS icon is displayed at the origin
#     """
#     acad.ActiveDocument.ActiveViewport.UCSIconOn = booleanOfUCSIcon
#     acad.ActiveDocument.ActiveViewport.UCSIconAtOrigin = booleanOfUCSatOrigin


# """
# Text

# Text Style Object
#     (1)SetFont method
#     object.SetFont(TypeFace, Bold, Italic, CharSet, PitchandFamily)
#     Function->Set the font for created text style object
#     object:text style object
#     TypeFace:string, font name, such as "宋体"
#     Bold:boolean, if True, bold, if False, normal
#     Italic:boolean, if True, italic, if False, normal
#     CharSet: long integer, defining font character set, the constants's meaning is as below
#         Constant             Meaning
#         0                    ANSI character set
#         1                     Default character set
#         2                     Symbol set
#         128                 Japanese character set
#         255                    OEM character set 
#     PitchandFamily: consists of 2 part:(a)Pitch, defining character's pitch(b)Family, defining character'stroke
#         Pitch:
#             Constant                     Meanning
#             0                             Default value
#             1                             Fixed value
#             2                             variable value
#         Family:
#             Conatant                     Meanning
#             0                             No consideration of stroke form
#             16                            Variable stroke width, with serif
#             32                             Variable stroke width, without serif
#             48                             Fixed stroke width, with or without serif
#             64                             Grass writting
#             80                             Old English stroke
#     (2) FontFile property
#     obj.fontFile = path
#     obj:textstyle object
#     set the given textstyle's font file by the path of character file, 
#     for example, path = acad.Path+r'\tssdeng.shx'

#     (3)BigFontFile property
#     obj.BigFontFile = path
#     obj:textstyle object
#     This property is similar to the FontFile property, except that it is used to specify 
#     an Asian-language Big Font file. The only valid file type is SHX

# """

# def CreateTextStyle(textStyleName):
#     """
#     testStyleName:string, the name of new created text style
#     """
#     return acad.ActiveDocument.TextStyles.Add(textStyleName)

# def ActivateTextStyle(textStyleObj):
#     """
#     Activate the created textstyle object
#     textStyleObj:textStyle object
#     """
#     acad.ActiveDocument.ActiveTextStyle = textStyleObj

# def GetActiveFontInfo():
#     """
#     return a tuple (typeFace, Bold, Italic, charSet, PitchandFamily) of the active textstyle object
#     """
#     return acad.ActiveDocument.ActiveTextStyle.GetFont()

# def SetActiveFontFile(path):
#     """
#     set the active textstyle's font file by the path of character file, 
#     for example, path = acad.Path+r'\tssdeng.shx'
#     """
#     acad.ActiveDocument.ActiveTextStyle.fontFile = path 
# def SetActiveBigFontFile(path):
#     """
#     This property is similar to the FontFile property, except that it is used to specify 
#     an Asian-language Big Font file. The only valid file type is SHX
#     """
#     acad.ActiveDocument.ActiveTextStyle.BigFontFile = path

# """
# Single Text

# Formatted text
#     (1)Alignment
#         object.Alignment = win32com.client.constants.X
#         [object.TextAlignmentPoint = pnt1]
#         [object.InsertionPoint = pnt2]
#         object:single text object
#         X:acAlignmentLeft 
#         acAlignmentCenter 
#         acAlignmentRight 
#         acAlignmentAligned 
#         acAlignmentMiddle 
#         acAlignmentFit 
#         acAlignmentTopLeft 
#         acAlignmentTopCenter 
#         acAlignmentTopRight 
#         acAlignmentMiddleLeft 
#         acAlignmentMiddleCenter 
#         acAlignmentMiddleRight 
#         acAlignmentBottomLeft 
#         acAlignmentBottomCenter 
#         acAlignmentBottomRight
#     Note that:Alignment property has to be set before TextAlignmentPoint or InsertionPoint property be set!
#     Text aligned to acAlignmentLeft uses the InsertionPoint property to position the text. Text aligned to 
#     acAlignmentAligned or acAlignmentFit uses both the InsertionPoint and TextAlignmentPoint properties to
#     position the text. Text aligned to any other position uses the TextAlignmentPoint property to position the text.

#     (2)InsertionPoint
#         object.InsertionPoint = pnt
#         pnt:Apoint type
#     Note:This property is read-only except for text whose Alignment property is set to acAlignmentLeft, 
#     acAlignmentAligned, or acAlignmentFit. To position text whose justification is other than left, aligned, 
#     or fit, use the TextAlignmentPoint property.

#     (3)ObliqueAngle
#         object.ObliqueAngle = rad
#         rad:Double, rad angle
#         The angle in radians within the range of -85 to +85 degrees. A positive angle denotes a lean to the right; 
#         a negative value will have 2*PI added to it to convert it to its positive equivalent. 

#     (4)Rotation
#         object.Rotation = rad
#         rad:Double, The rotation angle in radians. 

#     (5)TextAlignmentPoint
#     objcet.TextAlignmentPoint = pnt
#     pnt: Apoint type
#     Specifies the alignment point for text and attributes;Note that:Alignment property has to be set before 
#     TextAlignmentPoint or InsertionPoint property be set!Text aligned to acAlignmentLeft uses the InsertionPoint 
#     property to position the text.

#     (6)TextGenerationFlag
#     object.TextGenerationFlat = win32com.client.constants.x
#     X:acTextFlagBackward, acTextFlagUpsideDown
#     Specifies the attribute text generation flag, To specify both flags, add them together, 
#     that is acTextFlagBackward+acTextFlagUpsideDown

#     (7)TextString
#     object.TextString
#     This method returns the text string of single text object

#     (8)commen editing method:
#     ArrayPolar, ArrayRectangular, Copy, Delete, Mirror, Move, Rotate.
# """

# def AddText(insertPnt, height, textString):
#     """
#     add single text
#     textString:string, the inserted single text
#     insertPnt:Apoint type, insert point
#     height:the text height
#     """
#     return space.AddText(textString, Apoint(*insertPnt), height)

# """
# MutiText
# """
# def AddMText(textString, insertPnt, width):
#     """
#     Creates an MText entity in a rectangle defined by the insertion point and width of the bounding box.
#     textString:string
#     insertPnt:Apoint type
#     width:float, The width of the MText bounding box
#     """
#     return space.AddMText(Apoint(*insertPnt), width, textString)


# """
# Dimension and Tolerance

# Common property of dim object
#     (1)obj.DecimalSeparator = X
#     X:string, such as '.', can be any string.

#     (2)obj.ArrowheadSize = X
#     X:Double, The size of the arrowhead must be specified as a positive real > =  0.0, The initial value for this property is 0.1800.

#     (3)obj.DimensionLineColor = X
#     X:Use a color index number from 0 to 256, or bilt-in constants

#     (4)obj.DimLineInside = X
#     X:Boolean, default is False. Specifies the display of dimension lines inside the extension lines . Dimension line is the line below
#     the dimenion text and extension lines are a pair of lines pointing to the limit point of a dimension.

#     (5)obj.Fit = win32com.client.constants.X
#     Specifies the placement of text and arrowheads inside or outside extension lines, based on the available space between the extension lines
#     X:acTextAndArrows, acArrowsOnly, acTextOnly, acBestFit

#     (6)obj.Measurement
#     Read-only, returns the actural dimension value.

#     (7)obj.TextColor
#     the text color

#     (8)obj.TextHeight
#     the text height

#     (9)obj.TextOverride = X
#     X:string. ''represents the actural measurement.'<>'represents the actural measurment value, such as '<>mm'

#     (10)obj.Arrowhead1Type = win32com.client.constants.X
#     obj.Arrowhead2Type = win32com.client.constants.X
#     X:
#         acArrowDefault, acArrowDot, acArrowDotSmall, acArrowDotBlank, acArrowOpen, acArrowOblique, acArrowArchTick, etc.

#     (11)obj.TextPosition = X
#     X:Apoint type. the position of text.

#     (12)obj.TextPrefix = X
#     X:string

#     (13)obj.TextSuffix = X
#     X:string
#     (14)obj.UnitsFormat = win32com.client.constants.X
#     Specifies the unit format for all dimensions except angular
#     X:
#         acDimLScientific, acDimLDecimal, acDimLEngineering, acDimLArchitectural, acDimLFractional
#     The initial value for this property is acDimLDecimal.If this property is set to acDimLDecimal, 
#     the format specified by the DecimalSeparator and PrimaryUnitsPrecision properties will be used to format the decimal value

#     (15)obj.PrimaryUnitsPrecision = win32com.client.constants.X
#     Specifies the number of decimal places displayed for the primary units of a dimension or tolerance
#     X:
#         acDimPrecisionZero: 0
#         acDimPrecisionOne: 0.0
#         acDimPrecisionTwo: 0.00
#         acDimPrecisionThree: 0.000
#         acDimPrecisionFour: 0.0000 
#         acDimPrecisionFive: 0.00000
#         acDimPrecisionSix: 0.000000
#         acDimPrecisionSeven: 0.0000000
#         acDimPrecisionEight: 0.00000000 

#     (16)obj.VerticalTextPosition = win32com.client.constants.X
#     X:
#         acAbove, acOutside, acVertCentered, acJI

#     (17)obj.TextOutsideAlign = X
#     obj:
#     X:Boolean, Specifies the position of dimension text outside the extension lines for all dimension types except ordinate
#     True: Align the text horizontally
#     False: Align the text with the dimension line

#     (18)obj.CenterType = win32com.client.constants.X
#     Specifies the type of center mark for radial and diameter dimensions
#     obj: DimDiametric, DimRadial, DimRadialLarge 
#         X:
#         acCenterMark 
#         acCenterLine 
#         acCenterNone
#     Note:The center mark is visible only if you place the dimension line outside the circle or arc.

#     (19) obj.CenterMarkSize = X
#     Specifies the size of the center mark for radial and diameter dimensions.
#     X:Double, A positive real number specifying the size of the center mark or lines
#     Note:The initial value for this property is 0.0900. This property is not available if the CenterType property is set to acCenterNone.

#     (20)obj.ForceLineInside = X
#     Specifies whether a dimension line is drawn between the extension lines even when the text is placed outside the extension lines
#     X:Boolean
#     True: Draw dimension lines between the measured points when arrowheads are placed outside the measured points. 
#     False: Do not draw dimension lines between the measured points when arrowheads are placed outside the measured points

#     (21)obj.StyleName = X
#     X:string
#     Specifies the name of the style used with the object

# """

# def AddDimAligned(extPnt1, extPnt2, extPnt3):
#     """
#     Creates an aligned dimension object
#     extPnt1:Apoint type, the 3D WCS coordinates specifying the first endpoint of the extension line
#     extPnt2:Apoint type, the 3D WCS coordinates specifying the second endpoint of the extension line
#     textPosition:Apoint type, the 3D WCS coordinates specifying the text position
#     """
#     return space.AddDimAligned(Apoint(*extPnt1), Apoint(*extPnt2), Apoint(*extPnt3))

# def AddDimRotated(xlPnt1, xlPnt2, dimLineLocation, rotAngle):
#     """
#     Creates a rotated linear dimension
#     xlPnt1:Apoint type, the 3D WCS coordinates specifying the first endpoint of the extension line
#     xlPnt2:Apoint type, the 3D WCS coordinates specifying the first endpoint of the extension line
#     rotAngle:Double, The angle, in radians, of rotation displaying the linear dimension
#     """
#     return space.AddDimRotated(Apoint(*xlPnt1), Apoint(*xlPnt2), Apoint(*dimLineLocation), rotAngle)


# def AddDimLinear(pt1, pt2, direct="+x", dle=50):
#     x1, y1 = pt1[0], pt1[1]
#     x2, y2 = pt2[0], pt2[1]
#     try: z1 = pt1[2] 
#     except: pass
#     try: z2 = pt2[2] 
#     except: pass
#     min_x, min_y, max_x, max_y = min(x1, x2), min(y1, y2), max(x1, x2), max(y1, y2)

#     if dle < 0: raise ValueError("dle must >= 0 ... ... ")

#     if direct == "+x" :
#         pt3 = [max_x+dle, min_y]
#         AddDimRotated(pt1, pt2, pt3, math.radians(90))
#     if direct == "-x" :
#         pt3 = [min_x-dle, min_y]
#         AddDimRotated(pt1, pt2, pt3, math.radians(90))
#     if direct == "+y" :
#         pt3 = [min_x, max_y+dle]
#         AddDimRotated(pt1, pt2, pt3, math.radians(0))
#     if direct == "-y" :
#         pt3 = [max_x, min_y-dle]
#         AddDimRotated(pt1, pt2, pt3, math.radians(0))

#     if direct != "+x" and direct != "-x" and direct != "+y" and direct != "-y": raise ValueError("DimLinear::direct::must be +x or -x or +y or -y , not {direct}")


# def AddDimRadial(center, chordPnt, leaderLength):
#     """
#     Creates a radial dimension for the selected object at the given location
#     center:Apoint tyoe
#     chordPnt:Apoint type, The 3D WCS coordinates specifying the point on the circle or arc to attach the leader line
#     leaderLength:double, The positive value representing the length from the ChordPoint to the annotation text or dogleg
#     """
#     return space.AddDimRadial(Apoint(*center), Apoint(*chordPnt), leaderLength)



    
# def AddDimDiameter(center, angle, diameter):
#     rad = AngletoRad(angle)
#     xdr = math.cos(rad)
#     ydr = math.sin(rad)
#     direct = Vec3ResetLength([xdr, ydr], diameter/2)
#     direct2 = [-direct[0], -direct[1], -direct[2]]
#     pt2 = Vec3Add(center, direct)
#     pt1 = Vec3Add(center, direct2)
#     AddDimDiametric(pt2, pt1, rad)


# def AddDimDiametric(chordPnt, farChordPnt, leaderLength):
#     """
#     Creates a diametric dimension for a circle or arc given the two points on the diameter and the length of the leader line
#     chordPnt:Apoint type, The 3D WCS coordinates specifying the first diameter point on the circle or arc
#     farChordPnt:Apoint type, The 3D WCS coordinates specifying the second diameter point on the circle or arc
#     leaderLength:The positive value representing the length from the ChordPoint to the annotation text or dogleg, when it is 0, 
#     from obj.Fit = win32com.client.constants.acTextAndArrows can make the arrow and text inside the circle

#     """
#     return space.AddDimDiametric(Apoint(*chordPnt), Apoint(*farChordPnt), leaderLength)

# def AddDimAngular(vertex, firstPnt, secondPnt, textPnt):
#     """
#     Creates an angular dimension for an arc, two lines, or a circle
#     vertex, Apoint type, The 3D WCS coordinates specifying the center of the circle or arc, or the common vertex between the two dimensioned lines
#     firstPnt, Apoint type, The 3D WCS coordinates specifying the point through which the first extension line passes
#     secondPnt, Apoint type, The 3D WCS coordinates specifying the point through which the second extension line passes
#     textPnt, Apoint type, The 3D WCS coordinates specifying the point at which the dimension text is to be displayed
#     """
#     return space.AddDimAngular(Apoint(*vertex), Apoint(*firstPnt), Apoint(*secondPnt), Apoint(*textPnt))

# def AddDimOrdinate(definitionPnt, leaderPnt, axis):
#     """
#     Creates an ordinate dimension given the definition point and the leader endpoint
#     definitionPnt, Apoint type, The 3D WCS coordinates specifying the point to be dimensioned
#     leaderPnt, Apoint type, The 3D WCS coordinates specifying the endpoint of the leader. This will be the location at which the dimension text is displayed
#     axis, Boolean, True: Creates an ordinate dimension displaying the X axis value;False: Creates an ordinate dimension displaying the Y axis value
#     """
#     return space.AddDimOrdinate(Apoint(*definitionPnt), Apoint(*leaderPnt), axis)

# def AddLeader(pntArray, annotation = None, type = None):
#     """
#     Creates a leader line based on the provided coordinates or adds a new leader cluster to the MLeader object
#     pntArray, The array of 3D WCS coordinates, such as (1, 2, 3, 4, 5, 6), specifying the leader. You must provide at least two points to define the leader. The third point is optional
#     annotation, BlockReference, MText, Tolerance type.The object that should be attached to the leader. The value can also be NULL to not attach an 
#     Type:built-in contants, win32com.client.constants.X, X is as the following:
#     acLineNoArrow 
#     acLineWithArrow 
#     acSplineNoArrow 
#     acSplineWithArrow
#     >>>ann = acad.AddMText('demo', Apoint(30, 30, 0), 2)
#     >>>import win32com.client
#     >>>acad.AddLeader(0, 0, 0, 30, 30, 0, annotation = a, type = win32com.client.constants.acLineWithArrow)
#     """
#     return space.AddLeader(VtVertex(*pntArray), annotation, type)



# """
# Dimension style object

# (1)obj.CopyFrom(X)
# X:self.DimStyle0, self.ActiveDimStyle, and other dimension style object

# """
# def CreateDimStyle(name):
#     """
#     creat a new dimension style named name.
#     name:string
#     Once created, from the created dimstyle object's CopyFrom() method to get an existed dimstyle's
#     attributes.DimStyle0 is prefered to be copied from, when a new dim factor is needed, just reset corresponding system variable.For example:
#     >>acad = Autocad()
#     >>testDim = acad.CreateDimStyle('test')
#     >>acad.SetSystemVariable('dimlfac', 100)
#     >>testDim.CopyFrom(acad.DimStyle0)
#     """
#     return acad.ActiveDocument.DimStyles.Add(name)

# def DimStyleNumbers():
#     """
#     :return: int, the total number of defined dim style
#     """
#     return DimStyles().Count

# def DimStyleNames():
#     """
#     :return: list, all names of defined dim style
#     """
#     dimnames = []
#     for i in range(DimStyleNumbers()):
#         dimnames.append(DimStyles().Item(i).Name)
#     return dimnames
 
# def DimStyle0():
#     """
#     return created dimension style object whose index is 0 in modelspace, 

#     """
#     return space(0)
 
# def DimStyles():
#     """
#     return dimstyles object
#     """
#     return acad.ActiveDocument.DimStyles
 
# def ActiveDimStyle():
#     """
#     return a dim style set by system variable
#     """
#     return acad.ActiveDocument.ActiveDimStyle

# def GetDimStyle(dimname):
#     """
#     :param dimname:str or int, the name or index of dim style
#     :return:specified dim style object
#     """
#     if isinstance(dimname, str):
#         index = DimStyleNames().index(dimname)
#     elif isinstance(dimname, int):
#         index = dimname
#     else:
#         raise PycomError('dimname in GetDimStyle is wrong')
#     return DimStyles().Item(index)

# def SetActiveDimStyle(dimname):
#     """
#     Activate DimStyle dimname
#     :param dimname: str or int, the name or index of dim style
#     :return: None.
#     """
#     acad.ActiveDocument.ActiveDimStyle = GetDimStyle(dimname)


# def ChangeDimStyle(dimname):
#     SetActiveDimStyle(dimname)


# """
# Utility object method
# """
# def GetString(hasSpaces, Prompt = ''):
#     """
#     hasSpaces:
#         0:input string shall not has empty char('') meaning input has been done;
#         1:input string can have empty char(''), and the 'Entery' keystroke means the input process has been done.
#     Prompt:
#         string, default to None
#     """
#     return acad.ActiveDocument.Utility.GetString(hasSpaces, Prompt)

# def AngleFromXAxis(pnt1, pnt2):
#     """
#     Gets the angle of a line from the X axis
#     pnt1:Apoint type, The start point of the line;
#     pnt2:Apoint type, The endpoint of the line.
#     """
#     return acad.ActiveDocument.Utility.AngleFromXAxis(pnt1, pnt2)

# def GetAngle(basePnt=[0, 0, 0], prompt=''):
#     """
#     Gets the angle specified
#     """
#     return acad.ActiveDocument.Utility.GetAngle(Point=Apoint(*basePnt), Prompt=prompt)

# def GetPoint(Point=None, Prompt=''):
#     """
#     Gets the selected point
#     """
#     if Point:
#         return acad.ActiveDocument.Utility.GetPoint(Point=Apoint(*Point), Prompt=Prompt)
#     else:
#         return acad.ActiveDocument.Utility.GetPoint(Prompt=Prompt)
    
# def GetDist(pnt = '', prompt = ''):
#     """
#     Gets the point selected in AutoCAD
#     pnt:Apoint type, optional, The Point parameter specifies a relative base point in the WCS
#     """
#     if not pnt:
#         return acad.ActiveDocument.Utility.GetDistance(ArrayTransform(GetPoint()), prompt)
#     else:
#         return acad.ActiveDocument.Utility.GetDistance(pnt, prompt)
    
# def InitializeUserInput(bits, keywords):
#     """
#     Before from GetKeyword method, this method has to be used to limit the user-input forms , and this method
#     can also used with GetAngle, GetCorner, GetDistance, GetInteger, GetOrientation, GetPoint, GetReal, and cannot be
#     used with GetString.Unless it is set again, or it will control the type of input forever.
#     bits:integer
#         1: Disallows NULL input. This prevents the user from responding to the request by entering only [Return] or a space. 
#         2: Disallows input of zero (0). This prevents the user from responding to the request by entering 0. 
#         4: Disallows negative values. This prevents the user from responding to the request by entering a negative value. 
#         8: Does not check drawing limits, even if the LIMCHECK system variable is on. This enables the user to enter a point outside the current drawing limits. This condition applies to the next user-input function even if the AutoCAD LIMCHECK system variable is currently set. 
#         16: Not currently used. 
#         32: Uses dashed lines when drawing rubber-band lines or boxes. This causes the rubber-band line or box that AutoCAD displays to be dashed instead of solid, for those methods that let the user specify a point by selecting a location on the graphics screen. (Some display drivers use a distinctive color instead of dashed lines.) If the AutoCAD POPUPS system variable is 0, AutoCAD ignores this bit. 
#         64: Ignores Z coordinate of 3D points (GetDistance method only). This option ignores the Z coordinate of 3D points returned by the GetDistance method, so an application can ensure this function returns a 2D distance. 
#         128: Allows arbitrary input—whatever the user types. 
#     keywords:strings, such as 'width length height'
#     """
#     acad.ActiveDocument.Utility.InitializeUserInput(bits, keywords)

# def GetKeyword(prompt = ''):
#     """
#     Before from GetKeyword method, this method has to be used to limit the user-input forms
#     Gets a keyword string from the user
#     """
#     return acad.ActiveDocument.Utility.GetKeyword(prompt)

# def GetEntity():
#     """
#     Return a tuple containing the picked object and the coordinate of  picked point
#     """
#     return acad.ActiveDocument.Utility.GetEntity()

# def GetReal(prompt = ''):
#     """
#     Gets a real (double) value from the user.
#     """
#     return acad.ActiveDocument.Utility.GetReal(prompt)

# def GetInt(prompt = ''):
#     """
#     Gets an integer value from the user.
#     """
#     return acad.ActiveDocument.Utility.GetInteger(prompt)

# def Prompt(message):
#     """
#     Displays a prompt on the command line.
#     message:
#         string.
#     """
#     acad.ActiveDocument.Utility.Prompt(message)

# """
# Preferences object
# There are 9 sub-objects of preferences:
#     (1) Display--->acad.Preferences.Display
#     (2)Drafting--->acad.Preferences.Drafting
#     (3)Files--->acad.Preferences.Files
#     (4)OpenSave--->acad.Preferences.OpenSave
#     (5)Output--->acad.Preferences.Output
#     (6)Profiles--->acad.Preferences.Profiles
#     (7)Selection--->acad.Preferences.Selection
#     (8)System--->acad.Preferences.System
#     (9)User--->acad.Preferences.User

# """

# def Preferences():
#     """
#     return preferences object
#     """
#     return acad.Preferences




# # import math
# # XLine1Point = APoint(5, 25)
# # XLine2Point = APoint(25, 35)
# # DimLineLocation = APoint(10, 20)
# # RotationAngle = math.radians(0)
# # dimRotObj = acad.model.AddDimRotated(XLine1Point, XLine2Point, DimLineLocation, RotationAngle)
# # # XLine1Point 第一尺寸界线的起点；
# # # XLine2Point 第二尺寸界线的起点；
# # # DimLineLocation 尺寸线定位点，尺寸线或其延长线过该点；
# # # RotationAngle 尺寸线与水平方向的夹角，去弧度制；
# # # RotationAngle=0 水平标注，RotationAngle=90 竖直标注。




# # pyacad.ActiveDocument.SetVariable("PDMODE", 35) # 设置点样式的显示方式
# # pyacad.ActiveDocument.SetVariable("PDSIZE", 2) # 设置点大小





# if __name__ == '__main__':
#     table = {'dimclrd':62, 'dimdlI':0, 'dimclre':62, 
#         'dimexe':2, 'dimexo':3, 
#         'dimfxlon':1, 
#         'dimfxl':3, 'dimblk1':'_archtick', 
#         'dimldrblk':'_dot', 
#         'dimcen':2.5, 'dimclrt':62, 'dimtxt':3, 'dimtix':1, 
#         'dimdsep':'.', 'dimlfac':50}
    
#     AutoCad() 

#     if not IsEarlyBind: #判断是否是EarlyBind，如果不是则打开Earlybind模式
#         TurnOnEarlyBind() 
#     # acad.ShowLineweight(True)   #设置打开线宽
#     #进行绘制
#     line = AddLine([0, 0],  [100, 0]) #绘制线
#     circle = AddCircle([100, 0], 10)  #绘制圆
#     circleBig = AddCircle([0, 0], 110)
#     circleInner = AddCircle([0, 0], 90)
#     for i in range(3):
#         rad = math.radians(24)
#         Copy(line)
#         Copy(circle)
#         Rotate(line, [0, 0], rad)  
#         Rotate(circle, [0, 0], rad)


#     text = AddText( [0, 0], 20, 'Code makes a better world!')  #绘制文字
#     # erroe # text.Alignment = win32com.client.constants.acAlignmentTopCenter  #设置文字的alignment方向
#     Move(text, [0, 0], [0, -150]) #移动文字
#     underLine = AddLine([-175, -180], [175, -180]) #绘制下划线
#     # underLine.Lineweight = win32com.client.constants.acLnWt030 
#     Copy(underLine)
#     Offset(underLine, 5) #向上偏移下划线
#     # 保存文件
#     # SaveAsFile(r'pycomcad.dwg')
#     print("susess")




# """