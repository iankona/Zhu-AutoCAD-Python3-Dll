import math
import copy

import clr

import System

from Autodesk.AutoCAD.ApplicationServices import Application
from Autodesk.AutoCAD.EditorInput import PromptStatus, SelectionFilter, PromptSelectionOptions, SelectionSet, PromptIntegerOptions, PromptPointOptions, PromptDoubleOptions
# from Autodesk.AutoCAD.Runtime
from Autodesk.AutoCAD.DatabaseServices import Line, ObjectId, Transaction, OpenMode, BlockTable, BlockTableRecord, LayerTableRecord, ObjectIdCollection, TypedValue, DxfCode, DwgVersion
from Autodesk.AutoCAD.DatabaseServices import Extents3d, Polyline, Polyline3d, Line, Circle, Poly3dType, DBText, Region

from Autodesk.AutoCAD.Geometry import Point2d, Point3d, Point3dCollection, Matrix3d, Vector3d
# from Autodesk.AutoCAD.Internal.Utils import EntLast
from Autodesk.AutoCAD.Colors import Color, ColorMethod


# doc = Application.DocumentManager.MdiActiveDocument
# doc.Editor.WriteMessage("Hello, python, 鹅鹅鹅")
# ed = Application.DocumentManager.MdiActiveDocument.Editor


doc = None
ed = None
db = None
Command = None
CommandAsync = None


# def decorator(func):
#     def wrapper(*args, **kwargs):
#         print("befor")
#         result = func(*args, **kwargs)
#         print("after")
#         return result
#     return wrapper


class lock:
    def __enter__(self):
        GetActiveDocument()
        self.dlock = doc.LockDocument()
        return self.dlock
    def __exit__(self, type, value, traceback):
        self.dlock.Dispose()

trans = None
layertable = None
blocktable = None
modelblock = None
currentblock = None

class transaction:
    def __enter__(self):
        global trans, layertable, blocktable, modelblock, currentblock
        GetActiveDocument()
        self.trans = db.TransactionManager.StartTransaction() 
        self.layertable = self.trans.GetObject(db.LayerTableId, OpenMode.ForRead)
        self.blocktable = self.trans.GetObject(db.BlockTableId, OpenMode.ForRead)        
        model_space_objid = self.blocktable[BlockTableRecord.ModelSpace]
        self.modelblock = self.trans.GetObject(model_space_objid, OpenMode.ForWrite)
        self.currentblock = self.trans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite)
        trans, layertable, blocktable, modelblock, currentblock = self.trans, self.layertable, self.blocktable, self.modelblock, self.currentblock
        return self.trans
    def __exit__(self, type, value, traceback):
        self.trans.Commit()
        self.trans.Dispose()


class command_undo:
    def __enter__(self):
        GetUndo()
    def __exit__(self, type, value, traceback):
        SetUndo()


class command_osmode:
    def __enter__(self):
        GetOSMODE()
    def __exit__(self, type, value, traceback):
        SetOSMODE()


class test_context1:
    def __enter__(self):
        string1 = "管理器1.1之前...\n"
        print(string1), Prompt(string1)
    def __exit__(self, type, value, traceback):
        string2 = "管理器1.2之后...\n"
        print(string2), Prompt(string2)

class test_context2:
    def __enter__(self):
        string1 = "管理器2.1之前...\n"
        print(string1), Prompt(string1)
    def __exit__(self, type, value, traceback):
        string2 = "管理器2.2之后...\n"
        print(string2), Prompt(string2)


def decorator_command(func):
    def wrapper():
        GetActiveDocument()
        try:
            func()
        except Exception as e:
            print(e)
            # print(f"错误信息: {e}")
            # print(f"发生错误的文件: {e.__traceback__.tb_frame.f_globals['__file__']}")
            # print(f"发生错误的行号: {e.__traceback__.tb_lineno}")
            Prompt(f"...函数{func}出错...\n")
    return wrapper


def decorator_command_undo(func):
    def wrapper():
        GetActiveDocument()
        try:
            GetUndo()
            func()
            SetUndo()
        except Exception as e:
            HappenErrorUndo()
            print(e)
            Prompt(f"...函数{func}出错...\n")
    return wrapper



def decorator_db_lock(func):
    def wrapper(*args, **kwargs):
        GetActiveDocument()
        dlock = doc.LockDocument()
        result = func(*args, **kwargs)
        dlock.Dispose()
        return result
    return wrapper


def test_decorator1(func):
    def wrapper():
        GetActiveDocument()
        string1 = "装饰器1.1之前...\n"
        string2 = "装饰器1.2之后...\n"
        print(string1), Prompt(string1)
        result = func()
        print(string2), Prompt(string2)
        return result
    return wrapper
    
def test_decorator2(func):
    def wrapper():
        GetActiveDocument()
        string1 = "装饰器2.1之前...\n"
        string2 = "装饰器2.2之后...\n"
        print(string1), Prompt(string1)
        result = func()
        print(string2), Prompt(string2)
        return result
    return wrapper



def GetActiveDocument():
    global doc, ed, db, Command, CommandAsync
    doc = Application.DocumentManager.MdiActiveDocument
    ed = doc.Editor
    db = doc.Database
    Command = doc.Editor.Command
    CommandAsync = doc.Editor.CommandAsync
    return doc, ed, db, Command, CommandAsync




def ChangeObjectIdLayer(objid_list=[], layer_name="0"):
    AddLayer(layer_name)
    for objid in objid_list:
        objref = trans.GetObject(objid, OpenMode.ForWrite)
        objref.Layer = layer_name

def ChangeSelectionSetLayer(ss1:SelectionSet, layer_name="0"):
    AddLayer(layer_name)
    objid_list = ss1.GetObjectIds()
    for objid in objid_list:
        objref = trans.GetObject(objid, OpenMode.ForWrite)
        objref.Layer = layer_name

def AddLayer(layer_name="", color_index=0):
    if layertable.Has(layer_name) == False:
        layer_table_record = LayerTableRecord()
        layer_table_record.Name = layer_name
        if layer_name == "图层1": color_index = 33
        if layer_name == "打标1": color_index = 6
        if color_index != 0:
            layer_table_record.Color = Color.FromColorIndex(ColorMethod.ByAci, color_index)
        layertable.UpgradeOpen()
        layertable.Add(layer_table_record) # 将新图层追加到图层表
        layertable.DowngradeOpen()
        trans.AddNewlyCreatedDBObject(layer_table_record, True)



def CheckLayerAndColor(objref, layer_name, color_index):
    if layer_name != "": 
        AddLayer(layer_name)
        objref.Layer = layer_name
    if color_index != 0: 
        objref.Color = Color.FromColorIndex(ColorMethod.ByAci, color_index)

def AddLine(start_point, final_point, layer_name="", color_index=0):
    line = Line(ToPoint3d(start_point), ToPoint3d(final_point))
    CheckLayerAndColor(line, layer_name, color_index)
    currentblock.AppendEntity(line)
    trans.AddNewlyCreatedDBObject(line, True)
    return line


def AddLWPolyLine(ptlist, layer_name="", color_index=0):
    pline = Polyline()
    for i, pt1 in enumerate(ptlist):
        pline.AddVertexAt(i, ToPoint2d(pt1), 0, 0, 0)
    CheckLayerAndColor(pline, layer_name, color_index)
    currentblock.AppendEntity(pline)
    trans.AddNewlyCreatedDBObject(pline, True)
    return pline


def AddPoint():pass

def AddText(pt1, string="单行文字", size=50, layer_name="", color_index=0):
    text = DBText()
    text.Position = ToPoint3d(pt1) 
    text.TextString = string
    text.Height = size
    text.Rotation = 0
    # text.IsMirroredInX = True # 在X轴镜像
    # text.HorizontalMode = TextHorizontalMode.TextCenter # 设置对齐方式
    # text.AlignmentPoint = text.Position # 设置对齐点
    CheckLayerAndColor(text, layer_name, color_index)
    currentblock.AppendEntity(text)
    trans.AddNewlyCreatedDBObject(text, True)
    return text

def AddPolyline3d(ptlist, layer_name="", color_index=0):
    collection = Point3dCollection()
    for pt1 in ptlist:
        collection.Add(ToPoint3d(pt1))
    pline = Polyline3d(Poly3dType.SimplePoly, collection, False) # Closed = True
    CheckLayerAndColor(pline, layer_name, color_index)
    currentblock.AppendEntity(pline)
    trans.AddNewlyCreatedDBObject(pline, True)
    return pline

def GetObjectForRead(objid):
    objref = trans.GetObject(objid, OpenMode.ForRead)
    return objref

def GetObjectForWrite(objid):
    objref = trans.GetObject(objid, OpenMode.ForWrite)
    return objref




def Copy(objid, pt1, pt2, layer_name="", color_index=0):
    dr = Direct(pt1, pt2)
    vecdr = Vector3d(*dr)
    matrix4x4 = Matrix3d.Displacement(vecdr)
    objref = GetObjectForRead(objid)
    copyentity = objref.GetTransformedCopy(matrix4x4)
    CheckLayerAndColor(copyentity, layer_name, color_index)
    currentblock.AppendEntity(copyentity)
    trans.AddNewlyCreatedDBObject(copyentity, True)
    return copyentity


def Move(objid, pt1, pt2):
    dr = Direct(pt1, pt2)
    vecdr = Vector3d(*dr)
    matrix4x4 = Matrix3d.Displacement(vecdr)
    objref = GetObjectForWrite(objid)
    objref.TransformBy(matrix4x4)

def MoveCopy(objid, pt1, pt2, layer_name="", color_index=0):
    return Copy(objid, pt1, pt2, layer_name, color_index)



def Rotate3d(objid, angle=90, center=[0,0], axis=[0,0,1]):
    matrix4x4 = Matrix3d.Rotation(System.Double(angle), Vector3d(*axis), ToPoint3d(center))
    objref = GetObjectForWrite(objid)
    objref.TransformBy(matrix4x4)




def Rotate3dCopy(objid, angle=90, center=[0,0], axis=[0,0,1], layer_name="", color_index=0):
    matrix4x4 = Matrix3d.Rotation(System.Double(angle), Vector3d(*axis), ToPoint3d(center))
    objref = GetObjectForRead(objid)
    copyentity = objref.GetTransformedCopy(matrix4x4)
    CheckLayerAndColor(copyentity, layer_name, color_index)
    currentblock.AppendEntity(copyentity)
    trans.AddNewlyCreatedDBObject(copyentity, True)
    return copyentity



# class ZhuObject: pass
# entcalss = ZhuObject()
# entcalss.objid = None
# def __GetEntLastFunc():
#     result = ed.SelectLast() # PromptSelectionResult # (OK,[((1375515652304),NonGraphical,0,)])
#     ss1 = result.Value # SelectionSet (((2361431560400),NonGraphical,0,))
#     idlist = ss1.GetObjectIds()
#     entcalss.objid = idlist[0]




def EntLast():
    # CommandAddLine后，获取entlast偶尔会出现ss1为None的错误，但大部分对象大部分时间获取还是能用的。
    result = ed.SelectLast() # PromptSelectionResult # (OK,[((1375515652304),NonGraphical,0,)])
    ss1 = result.Value # SelectionSet (((2361431560400),NonGraphical,0,))
    if ss1 == None: raise ValueError(f"SelectionSet为None, 未能获取到entlast...")
    idlist = ss1.GetObjectIds()
    return idlist[0]


def EntLastSet():
    result = ed.SelectLast()
    ss1 = result.Value
    if ss1 == None: raise ValueError(f"SelectionSet为None, 未能获取到entlastset...")
    return ss1


def SelectionSetFromID(objid:ObjectId):
    return SelectionSet.FromObjectIds([objid])


def ToPoint2d(pt0):
    x, y = pt0[0:2]
    return Point2d(x, y)


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




def CommandCopy(objid:ObjectId|SelectionSet):
    Command(["copy", objid, "D", ""]), Prompt("\n")
    

def CommandCopyMove(objid:ObjectId|SelectionSet, start_point, final_point):
    Command(["copy", objid, ToPoint3d(start_point), ToPoint3d(final_point),"E"]), Prompt("\n")
    

def CommandMove(objid:ObjectId|SelectionSet, start_point, final_point):
    Command(["move", SelectionSetFromID(objid), "", ToPoint3d(start_point), ToPoint3d(final_point)]), Prompt("\n")
    

def CommandOffSet(objid:ObjectId|SelectionSet, distance, directpt1=[], 图层=""): # directpt1 = [x, y]
    if 图层 == "当前":
        Command(["OFFSET", "L", "C", System.Double(distance), "E"]) # 当前图层
        Command(["OFFSET", "", objid, ToPoint3d(directpt1), "E"])
        Command(["OFFSET", "L", "S", "", "E"]) # 还原回原图层
    else: 
        Command(["OFFSET", System.Double(distance), objid, ToPoint3d(directpt1), "E"]) # 若是SelectionSet，则只会偏移第1个objid  
    Prompt("\n")
    

def CommandRotate(objid:ObjectId|SelectionSet, centert_point, angle):
    Command(["rotate", objid, ToPoint3d(centert_point), System.Double(angle)]), Prompt("\n")
    

def CommandRotateCopy(objid:ObjectId|SelectionSet, centert_point, angle):
    Command(["rotate", objid, ToPoint3d(centert_point), "C", System.Double(angle)]), Prompt("\n")
    



def CommandErase(objid:ObjectId|SelectionSet):
    Command(["rotate", objid]), Prompt("\n")


def CommandChangeLayer(obj, layername):
    CommandAddLayer(layername)
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


def Vec3XYtoXZ(veclist=[]):
    result = []
    for vec in veclist:
        x, y, z = Vec2toVec3(vec) 
        result.append([x,z,y])
    return result
    

def Absolute(value):
    if isinstance(value, list):
        result = []
        for va in value:
            result.append(-va) if va < 0 else result.append(va)
        return result
    if value < 0: return -value
    return value


def Distance(pt1, pt2):
    x1, y1, z1 = Vec2toVec3(pt1)
    x2, y2, z2 = Vec2toVec3(pt2)
    a = (x2-x1)**2 + (y2-y1)**2 + (z2-z1)**2
    return math.sqrt(a)

def Direct(pt1, pt2):
    x1, y1, z1 = Vec2toVec3(pt1)
    x2, y2, z2 = Vec2toVec3(pt2)
    return [x2-x1, y2-y1, z2-z1]


def DirectListToPointList(pt, drlist):
    pt1 = pt
    result = [pt1]
    for dr1 in drlist:
        pt1 = Vec3Add(pt1, dr1)
        result.append(pt1)
    return result


def GetAttachNDirectPointList(pt1, pt2, length):
    dr1 = Direct(pt1, pt2)
    dr1 = Vec3ResetLength(dr1, length)
    dr2 = Direct(pt2, pt1)
    dr2 = Vec3ResetLength(dr2, length)    
    return [Vec3Add(pt1, dr1), Vec3Add(pt2, dr2)]


def GetAttachNDirectPointPt1(pt1, pt2, length):
    dr1 = Direct(pt1, pt2)
    dr1 = Vec3ResetLength(dr1, length) 
    return Vec3Add(pt1, dr1)


def GetAttachNDirectPointPt2(pt1, pt2, length):
    dr2 = Direct(pt2, pt1)
    dr2 = Vec3ResetLength(dr2, length)    
    return Vec3Add(pt2, dr2)


def GetAttachWDirectPointList(pt1, pt2, length):
    dr1 = Direct(pt2, pt1)
    dr1 = Vec3ResetLength(dr1, length)
    dr2 = Direct(pt1, pt2)
    dr2 = Vec3ResetLength(dr2, length)    
    return [Vec3Add(pt1, dr1), Vec3Add(pt2, dr2)]


def GetAttachWDirectPointPt1(pt1, pt2, length):
    dr1 = Direct(pt2, pt1)
    dr1 = Vec3ResetLength(dr1, length)
    return Vec3Add(pt1, dr1)

def GetAttachWDirectPointPt2(pt1, pt2, length):
    dr2 = Direct(pt1, pt2)
    dr2 = Vec3ResetLength(dr2, length)    
    return Vec3Add(pt2, dr2)

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


def CommandAddPoint(pt1):
    Command(["point", ToPoint3d(pt1)]), Prompt("\n")
    


def CommandAddText(pt1, string, size): # pt1 = [x, y, z]
    Command(["text", ToPoint3d(pt1), System.Int32(size), System.Int32(0), string]), Prompt("\n")
    


def CommandAddLine(pt1, pt2): # pt1 = [x, y, z]
    Command(["LINE", ToPoint3d(pt1), ToPoint3d(pt2), ""]), Prompt("\n")
    



# def CommandAsyncAddLine(pt1, pt2): # pt1 = [x, y, z] # Error 
#     entcalss.objid = None
#     casync, zhu = CommandAsync(["LINE", ToPoint3d(pt1), ToPoint3d(pt2), ""]), Prompt("\n")
#     casync.OnCompleted(System.Action(__GetEntLastFunc))
#     print("entcalss.objid = ", entcalss.objid)
#     return entcalss.objid
    

def CommandAddPLine(ptlist=[]): # 函数(pt1, pt2, pt3...)
    列表 = [ToPoint3d(pt1) for pt1 in ptlist]
    Command(["PLINE"] + 列表 + [""]), Prompt("\n")
    

def CommandAddRect(pt1, pt2):
    Command(["RECTANG", ToPoint3d(pt1), ToPoint3d(pt2), ""]), Prompt("\n")
    

def CommandAddCircle(pt1, radius):
    CommandAddCircleR(pt1, radius)

def CommandAddCircleR(pt1, radius):
    Command(["CIRCLE", ToPoint3d(pt1), System.Double(radius), ""]), Prompt("\n")
    

def CommandAddCircleD(pt1, diameter):
    Command(["CIRCLE", ToPoint3d(pt1), "D", System.Double(diameter), ""]), Prompt("\n")
    


def CommandAddCircle2P(pt1, pt2):
    Command(["CIRCLE", "2P", ToPoint3d(pt1), ToPoint3d(pt2), ""]), Prompt("\n")
    

def CommandAddCircle3P(pt1, pt2, pt3):
    Command(["CIRCLE", "3P", ToPoint3d(pt1), ToPoint3d(pt2), ToPoint3d(pt3), ""]), Prompt("\n")
    

def GetStartPoint(objid):
    objref = trans.GetObject(objid, OpenMode.ForRead)
    start = objref.StartPoint
    return [start.X, start.Y, start.Z]

def GetFinalPoint(objid):
    objref = trans.GetObject(objid, OpenMode.ForRead)
    final = objref.EndPoint
    return [final.X, final.Y, final.Z]


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
    if dxfcode_filter_list == []:
        # value = [TypedValue(System.Int32(0), "Circle")] # == AutoLisp (DxfCode . "Circle") 
        typevalue_list = []
        for [dxfcode, checkchar] in dxfcode_filter_list:
            typevalue_list.append(TypedValue(System.Int32(dxfcode), checkchar))
        # options = PromptSelectionOptions()
        # print(typevalue_list)
        filter = SelectionFilter(typevalue_list)
        result = ed.GetSelection(filter) # PromptSelectionResult
    else:
        result = ed.GetSelection()
    ss1 = result.Value # SelectionSet
    # ids = ss1.GetObjectIds()
    # ss2 = SelectionSet.FromObjectIds([ids[-1]])
    Highlight(ss1)
    return ss1 

def SSGetIdList(dxfcode_filter_list=[]):
    ss1 = SSGet(dxfcode_filter_list)
    if ss1 == None: return []
    return ss1.GetObjectIds()


def Highlight(ss1:SelectionSet):
    with transaction() as trans:
        for objid in ss1.GetObjectIds():
            objref = trans.GetObject(objid, OpenMode.ForRead)
            objref.Highlight()

def GetEntityBound(objid:ObjectId):
    with transaction() as trans:
        objref = trans.GetObject(objid, OpenMode.ForRead)
        extend = objref.GeometricExtents
        point1 = extend.MinPoint
        point2 = extend.MaxPoint
    return [point1.X, point1.Y, point1.Z], [point2.X, point2.Y, point2.Z]

def GetEntityBoundXY(objid:ObjectId):
    [x1,y1,z1],[x2,y2,z2] = GetEntityBound(objid)
    return [x1, y1, 0], [x2, y2, 0]

def GetEntityBoundCenterXY(ss1:SelectionSet):
    [x1,y1,z1],[x2,y2,z2] = GetEntityBoundXY(ss1)
    return [(x1+x2)/2, (y1+y2)/2, 0]


def GetBound(ss1:SelectionSet):
    with transaction() as trans:
        extend = Extents3d()
        for objid in ss1.GetObjectIds():
            objref = trans.GetObject(objid, OpenMode.ForRead)
            extend.CommandAddExtents(objref.GeometricExtents) 
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
    with transaction() as trans:
        pline = trans.GetObject(objid, OpenMode.ForRead)
        result = []
        for i in range(pline.NumberOfVertices):
            point = pline.GetPoint3dAt(i)
            result.append([point.X, point.Y, point.Z])
    if pline.Closed: return result + result[0:1]
    return result



def GetLWPolyLineStartMid(objid:ObjectId):
    pline = trans.GetObject(objid, OpenMode.ForRead)
    point1 = pline.GetPoint3dAt(0)
    point2 = pline.GetPoint3dAt(1)
    pt1 = [point1.X, point1.Y, point1.Z]
    pt2 = [point2.X, point2.Y, point2.Z]
    mid = MidPt1Pt2(pt1, pt2)
    return mid

def GetLWPolyLineDirectList(objid:ObjectId):
    ptlist = GetLWPolyLinePointList(objid)
    drlist = []
    for i in range(len(ptlist)-1):
        pt1 = ptlist[i]
        pt2 = ptlist[i+1]
        dr1 = Direct(pt1, pt2)
        drlist.append(dr1)
    return drlist

def GetLWPolyLineMidPointList(objid:ObjectId):
    pline_point_list = GetLWPolyLinePointList(objid)
    midptlist = []
    for i in range(len(pline_point_list)-1):
        pt1 = pline_point_list[i]
        pt2 = pline_point_list[i+1]
        midptlist.append(MidPt1Pt2(pt1, pt2))
    return midptlist




def ChangeCoordinateXY(drlist, target_coord1="-Y", target_coord2="X"):
    result = []
    for dr in drlist:
        match target_coord1:
            case "-X": x = -dr[0]
            case  "X": x =  dr[0]
            case "+X": x =  dr[0]
            case "-Y": x = -dr[1]
            case  "Y": x =  dr[1]
            case "+Y": x =  dr[1]
            case _: raise ValueError(f"...未支持的坐标参数{target_coord1}...")
        match target_coord2:
            case "-X": y = -dr[0]
            case  "X": y =  dr[0]
            case "+X": y =  dr[0]
            case "-Y": y = -dr[1]
            case  "Y": y =  dr[1]
            case "+Y": y =  dr[1]
            case _: raise ValueError(f"...未支持的坐标参数{target_coord2}...")
        z = dr[2]
        result.append([x ,y ,z])
    return result









def GetLWPolyLineNormal(objid:ObjectId):
    with transaction() as trans:
        pline = trans.GetObject(objid, OpenMode.ForRead)
        return pline.Normal





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

def GetPoint3(string1="", string2="", string3=""):
    pt1 = GetPoint(string1)
    if pt1 == None: return None, None, None
    pt2 = GetPoint(string2)
    if pt2 == None: return None, None, None
    pt3 = GetPoint(string3)
    if pt3 == None: return None, None, None
    return pt1, pt2, pt3


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
    return 列表

 

# (command "style" "新字体样式名称" "字体文件名称" 0 1 0 0 0 0)
# (command "selectall")
# (command "chprop" "style" "新字体样式名称" "")
def CommandAddFontStyle(style_name:str, font_name:str):
    Command(["-style", style_name, font_name, System.Int32(0), System.Int32(1), System.Int32(0), "N", "N"])


def CommandChangeFontStyle(style_name:str, new_font_name:str):
    CommandAddFontStyle(style_name, new_font_name)



# (command "-style" "mystyle" "txt.shx,gbcbig.shx" 8 1 0 "N" "N" "N" )
# (command "-style" "mystyle" "Times New Roman" 8 1 0 "N" "N")
# 这两个都能正确运行，但很明显，第一行最后是三个N，第二行才两个，但不知道为什么这样，可能是“语境”吧
def CommandChangeStandardFontStyle(new_font_name:str):
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





