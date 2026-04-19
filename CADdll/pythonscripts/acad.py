import math
import copy
import time
import clr

import System

from Autodesk.AutoCAD.ApplicationServices import Application
from Autodesk.AutoCAD.EditorInput import SelectionMethod, PromptStringOptions, SubtractedKeywords, AddedKeywords, PromptStatus, SelectionFilter, PromptSelectionOptions, SelectionSet, PromptIntegerOptions, PromptPointOptions, PromptDoubleOptions
from Autodesk.AutoCAD.EditorInput import SelectedObject, SelectionMethod, PromptNestedEntityOptions, PromptDistanceOptions

# from Autodesk.AutoCAD.Runtime
from Autodesk.AutoCAD.DatabaseServices import Line, Arc, ObjectId, Spline, Transaction, OpenMode, BlockTable, BlockTableRecord, BlockReference, LayerTableRecord, ObjectIdCollection, TypedValue, DxfCode, DwgVersion
from Autodesk.AutoCAD.DatabaseServices import Entity, DBPoint, Extents3d, Polyline, Polyline3d, Line, Circle, Poly3dType, DBText, MText, Region, DBObjectCollection, Intersect, Group
from Autodesk.AutoCAD.DatabaseServices import RotatedDimension, AlignedDimension, FullSubentityPath, AssocPersSubentityIdPE


from Autodesk.AutoCAD.Geometry import Point2d, Point3d, Point3dCollection, Matrix3d, Vector2d, Vector3d, DoubleCollection
from Autodesk.AutoCAD.Colors import Color, ColorMethod
from Autodesk.AutoCAD.Internal import Utils

from Autodesk.AutoCAD.DatabaseServices import SubentityId, SubentityType, FullSubentityPath, MeshFaceterData, IdMapping, ObjectIdCollection, SubDMesh, Curve, Extents3d, Polyline, Polyline3d, Line, Circle, Poly3dType, DBText, Region, BooleanOperationType
from Autodesk.AutoCAD.Geometry import Point2d, Point3d, Point3dCollection, Matrix3d, Vector3d, NurbCurve3d
import System
from Autodesk.AutoCAD.BoundaryRepresentation import PointContainment, Brep, Face, BrepEntity

from System.Reflection import Assembly

assemly = Assembly.LoadFile("E:\\AutoCAD\AutoCAD 2023\\Acdbmgd.dll")
AssocPersSubentityIdPEType = assemly.GetType("Autodesk.AutoCAD.DatabaseServices.AssocPersSubentityIdPE")
TypeObjectId = assemly.GetType("Autodesk.AutoCAD.DatabaseServices.ObjectId")

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
        self.old_trans, self.old_layertable, self.old_blocktable, self.old_modelblock, self.old_currentblock = trans, layertable, blocktable, modelblock, currentblock
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
        global trans, layertable, blocktable, modelblock, currentblock
        self.trans.Commit()
        self.trans.Dispose()
        trans, layertable, blocktable, modelblock, currentblock = self.old_trans, self.old_layertable, self.old_blocktable, self.old_modelblock, self.old_currentblock


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



def GetDistance(default_distance, string=""):
    if string == "": string = "请输入距离: "
    option = PromptDistanceOptions(string)
    option.DefaultValue = default_distance
    result = ed.GetDistance(option)
    Prompt("\n")
    if result.Status == PromptStatus.OK: return result.Value
    return None

def GetString(default_string:str, message=""):
    if message == "": message = "请输入字符串: "
    option = PromptStringOptions(message)
    option.DefaultValue = default_string
    result = ed.GetString(option) # PromptResult # (OK,4000x123) # [result.Status, result.StringResult]
    Prompt("\n")
    if result.StringResult == "": return None
    if result.Status == PromptStatus.OK: return result.StringResult
    return None


def GetInt(default_int:int, string=""):
    if string == "": string = "请输入整数: "
    options = PromptIntegerOptions(string)
    options.DefaultValue = default_int
    result = ed.GetInteger(options) # PromptResult # (OK,4000x123) # [result.Status, result.StringResult]
    Prompt("\n")
    if result.Status == PromptStatus.OK: return result.Value
    return None

# GetPoint: ((OK,),(225.037456173945,41.6177530027926,0))
# GetPoint: ((Cancel,),(0,0,0))
def GetPoint(string="", base_point=[]):
    if string == "": string = "请选择顶点: "
    options = PromptPointOptions(string)
    if base_point != []: 
        options.BasePoint = ToPoint3d(base_point)
        options.UseBasePoint = True
    result = ed.GetPoint(options)
    Prompt("\n")
    if result.Status == PromptStatus.OK: return [result.Value.X, result.Value.Y, result.Value.Z]
    return None


def GetPoint2(string1="", string2=""):
    if string1 == "": string1 = "请选择第1个点:"
    if string2 == "": string2 = "请选择第2个点:"
    pt1 = GetPoint(string1)
    if pt1 == None: return None, None
    pt2 = GetPoint(string2, pt1)
    if pt2 == None: return None, None
    return pt1, pt2


def GetPoint3(string1="", string2="", string3="", baseptflag=False):
    if string1 == "": string1 = "请选择第1个点:"
    if string2 == "": string2 = "请选择第2个点:"
    if string3 == "": string3 = "请选择第3个点:"
    pt1 = GetPoint(string1)
    if pt1 == None: return None, None, None
    if baseptflag:
        pt2 = GetPoint(string2, pt1)
    else:
        pt2 = GetPoint(string2)
    if pt2 == None: return None, None, None
    pt3 = GetPoint(string3)
    if pt3 == None: return None, None, None
    return pt1, pt2, pt3

def GetCorner(string, base_point):
    if string == "": string = "请选择顶点: "
    result = ed.GetCorner(string, ToPoint3d(base_point))
    if result.Status == PromptStatus.OK: return [result.Value.X, result.Value.Y, result.Value.Z]
    return None

def GetCorner2(string1="", string2=""):
    pt1 = GetPoint(string1)
    if pt1 == None: return None, None
    pt2 = GetCorner(string2, pt1)
    if pt2 == None: return None, None
    return [pt1, pt2]



# GetDouble: ((OK,),45) # 回车 or 右键
# GetDouble: ((Cancel,),0) # ESC
def GetDouble(default_double:float, string=""):
    if string == "": string = "请输入数值: "
    options = PromptDoubleOptions(string)
    options.DefaultValue = default_double
    result = ed.GetDouble(options) 
    Prompt("\n")
    if result.Status == PromptStatus.OK: return result.Value
    return None

def GetDoubleListLimitCount(count=100):
    列表 = []
    for i in range(count):
        result = ed.GetDouble(f"请输入第{i+1}个数据:") 
        if result.Value == 0: break
        if result.Status == PromptStatus.OK: 列表.append(result.Value)
        else: break
    return 列表

def GetSelectFence(pt1, pt2, dxfcode_filter_list=[]):
    collect = Point3dCollection()
    collect.Add(ToPoint3d(pt1))
    collect.Add(ToPoint3d(pt2))
    if dxfcode_filter_list != []:
        typevalue_list = []
        for [dxfcode, checkchar] in dxfcode_filter_list:
            typevalue_list.append(TypedValue(System.Int32(dxfcode), checkchar))
        filter = SelectionFilter(typevalue_list)
        result = ed.SelectFence(collect, filter)
    else:
        result = ed.SelectFence(collect)
    ss1 = result.Value
    return ss1

def GetSelectFenceIdList(pt1, pt2, dxfcode_filter_list=[]):
    ss1 = GetSelectFence(pt1, pt2, dxfcode_filter_list)
    if ss1 == None: return []
    return [objid for objid in ss1.GetObjectIds()]


def GetSelectCorner(pt1, pt2, dxfcode_filter_list=[]):  # 边界相交不会被选择
    if dxfcode_filter_list != []:
        typevalue_list = []
        for [dxfcode, checkchar] in dxfcode_filter_list:
            typevalue_list.append(TypedValue(System.Int32(dxfcode), checkchar))
        filter = SelectionFilter(typevalue_list)
        result = ed.SelectWindow(ToPoint3d(pt1), ToPoint3d(pt2), filter)
    else:
        result = ed.SelectWindow(ToPoint3d(pt1), ToPoint3d(pt2))
    ss1 = result.Value
    return ss1

def GetSelectCornerIdList(pt1, pt2, dxfcode_filter_list=[]):
    ss1 = GetSelectCorner(pt1, pt2, dxfcode_filter_list)
    if ss1 == None: return []
    return [objid for objid in ss1.GetObjectIds()]



def GetSelectCornerCross(pt1, pt2, dxfcode_filter_list=[]): # 边界相交会被选择
    if dxfcode_filter_list != []:
        typevalue_list = []
        for [dxfcode, checkchar] in dxfcode_filter_list:
            typevalue_list.append(TypedValue(System.Int32(dxfcode), checkchar))
        filter = SelectionFilter(typevalue_list)
        result = ed.SelectCrossingWindow(ToPoint3d(pt1), ToPoint3d(pt2), filter)
    else:
        result = ed.SelectCrossingWindow(ToPoint3d(pt1), ToPoint3d(pt2))
    ss1 = result.Value
    return ss1

def GetSelectCornerCrossIdList(pt1, pt2, dxfcode_filter_list=[]):
    ss1 = GetSelectCornerCross(pt1, pt2, dxfcode_filter_list)
    if ss1 == None: return []
    return [objid for objid in ss1.GetObjectIds()]



def GetSelectPick(pt1, dxfcode_filter_list=[]):
    ss1 = GetSelectCornerCross(pt1, pt1, dxfcode_filter_list)
    return ss1


def GetSelectPickId(pt1, dxfcode_filter_list=[]):
    ss1 = GetSelectCornerCross(pt1, pt1, dxfcode_filter_list)
    objidlist = ss1.GetObjectIds()
    return objidlist[0]

def EntLast():
    # CommandAddLine后，获取entlast偶尔会出现ss1为None的错误，但大部分对象大部分时间获取还是能用的。
    result = ed.SelectLast() # PromptSelectionResult # (OK,[((1375515652304),NonGraphical,0,)])
    ss1 = result.Value # SelectionSet (((2361431560400),NonGraphical,0,))
    str(ss1)
    if ss1 == None: raise ValueError(f"SelectionSet为None, 未能获取到entlast...")
    objidlist = ss1.GetObjectIds()
    return objidlist[0]

def EntLastSet():
    result = ed.SelectLast()
    ss1 = result.Value
    if ss1 == None: raise ValueError(f"SelectionSet为None, 未能获取到entlastset...")
    return ss1

def SelectionSetFromID(objid:ObjectId):
    return SelectionSet.FromObjectIds([objid])

def Prompt(string):
    ed.WriteMessage(str(string))

def EntSelEnt(string: str=""):
    return EntSelEntity(string)
def EntSelEntity(string: str=""):
    if string == "": string = "请选择对象: "
    result = ed.GetEntity(string)  # == AutoLisp entsel
    Prompt(f"[{result.PickedPoint}, (图元: {result.ObjectId})]\n")
    pickpt = [result.PickedPoint.X, result.PickedPoint.Y, result.PickedPoint.Z]
    objid = result.ObjectId
    return [pickpt, objid]

def EntSelSub(dxfcode_filter_list=[], string: str=""):
    # Erroer 001 
    # if string == "": string = "请选择对象: "
    # result = ed.GetEntity(string)  # == AutoLisp entsel
    # objid = result.ObjectId
    # pickpoint = result.PickedPoint
    # selectobject = SelectedObject(objid, SelectionMethod.SubEntity, System.IntPtr.Zero)
    # print(selectobject.GetSubentities())

    # # Erroer 002
    # option = PromptNestedEntityOptions("\nPick a nested entity:")
    # result = ed.GetNestedEntity(option)
    # ids = [objid for objid in result.GetContainers()][::-1] # 可以选择block里的圆弧等对象
    # target_id = result.ObjectId
    # result_ids = ids + [target_id]
    # subent = SubentityId(SubentityType.Null, System.IntPtr.Zero)
    # path = FullSubentityPath(result_ids, subent)
    # with transaction() as trans:
    #     subobjref = trans.GetObject(result_ids[0], OpenMode.ForRead);                      
    # subobjref.Highlight(path, False)
    # print(option, "\n", result, "\n", ids, "\n", target_id, "\n", result_ids, "\n", subent, "\n", path, "\n", subobjref)

    a = System.Array.CreateInstance(TypeObjectId, System.Int32(10)) # Autodesk.AutoCAD.DatabaseServices.ObjectId[]

    result = ed.GetEntity(string) 
    with transaction() as trans:
        entity = trans.GetObject(result.ObjectId, OpenMode.ForWrite)
    # path = entity.GetSubentityPathsAtGraphicsMarker(SubentityType.Edge, System.IntPtr.Zero, result.PickedPoint, Matrix3d.Identity, System.Array.CreateInstance(TypeObjectId, System.Int32(10))) # 
    path = entity.GetSubentityPathsAtGraphicsMarker(SubentityType.Edge, System.Int64(0), result.PickedPoint, Matrix3d.Identity, System.Int32(0), System.Array.CreateInstance(TypeObjectId, System.Int32(10))) # 
    
    entity.Highlight(path[0], highlightAll=False)




    # # Erroer 003
    # per = ed.GetEntity(string) 
    # with transaction() as trans:
    #     entity = trans.GetObject(per.ObjectId, OpenMode.ForRead)
    # entId = [entity.ObjectId]
    # pSubentityIdPE = entity.QueryX(AssocPersSubentityIdPE.GetClass(AssocPersSubentityIdPEType)) # 索引超出了数组界限。在 Autodesk.AutoCAD.Runtime.RXObject.GetClass(Type type)
    # # No method matches given arguments for RXObject.GetClass: (<class 'clr._internal.GCOffsetBase'>)
    # if pSubentityIdPE == System.IntPtr.Zero: return 
    # subentityIdPE = AssocPersSubentityIdPE.Creat(pSubentityIdPE, False)
    # vertexIds = subentityIdPE.GetAllSubentities(entity, SubentityType.Vertex)
    # for subentId in vertexIds:
    #     path = FullSubentityPath(entId, subentId)
    #     point = entity.GetSubentity(path) # DBPoint
    #     print(point)

    # edgeIds = subentityIdPE.GetAllSubentities(entity, SubentityType.Edge)
    # for subentId in edgeIds:
    #     path = FullSubentityPath(entId, subentId)
    #     edge = entity.GetSubentity(path) # Entity

    # faceIds = subentityIdPE.GetAllSubentities(entity, SubentityType.Face)
    # for subentId in faceIds:
    #     path = FullSubentityPath(entId, subentId)
    #     face = entity.GetSubentity(path) # Entity


def EntSel(dxfcode_filter_list=[], string: str=""):
    # if string == "": string = "请选择对象: "
    # result = ed.GetEntity(string)  # == AutoLisp entsel
    # Prompt(f"(图元: {result.ObjectId})\n")
    ss1 = SSGet(dxfcode_filter_list=dxfcode_filter_list, sel_method=":S", string=string)
    if ss1 == None: return None
    objidlist = ss1.GetObjectIds()
    return objidlist[0]

# def SSSetFirst(ss1:SelectionSet):
#     ids = ss1.GetObjectIds()
#     ed.SetImpliedSelection(ids)

# def SSIdsFirst(ids):
#     ed.SetImpliedSelection(ids)

def SSGet(dxfcode_filter_list=[], sel_method="", string="", sssetfirst=False): # [[0, "Circle"], [0, "Line"]]
    if string == "": string = "选择对象:"
    option = PromptSelectionOptions()
    option.MessageForAdding = string
    match sel_method:
        case ":D": option.AllowDuplicates = True
        case ":A": option.SinglePickInSpace = True
        case ":E": option.SelectEverythingInAperture = True
        case ":N": option.PrepareOptionalDetails = True
        case ":S": option.SingleOnly = True
        case ":U": option.AllowSubSelections = True
        case ":V": option.ForceSubSelections = True
        case ":L": option.RejectObjectsOnLockedLayers = True
        case ":C": option.RejectObjectsFromNonCurrentSpace = True
        case ":P": option.RejectPaperspaceViewport = True

        # case "+#": option.AddKeywordsToMinimalList(AddedKeywords.LastAllPrevious)
        # case "+.": option.AddKeywordsToMinimalList(AddedKeywords.PickImplied)
        # case "+A": option.AddKeywordsToMinimalList(AddedKeywords.All)
        # case "+B": option.AddKeywordsToMinimalList(AddedKeywords.WindowCrossingBoxWPolygonCPolygon)
        # case "+M": option.AddKeywordsToMinimalList(AddedKeywords.Multiple)
        # case "+C": option.AddKeywordsToMinimalList(AddedKeywords.CrossingCPolygon)
        # case "+F": option.AddKeywordsToMinimalList(AddedKeywords.Fence)
        # case "+L": option.AddKeywordsToMinimalList(AddedKeywords.Last)
        # case "+P": option.AddKeywordsToMinimalList(AddedKeywords.Previous)
        # case "+W": option.AddKeywordsToMinimalList(AddedKeywords.WindowWPolygon)
        # case "+G": option.AddKeywordsToMinimalList(AddedKeywords.Group)

        # case "-#": option.RemoveKeywordsFromFullList(SubtractedKeywords.LastAllGroupPrevious)
        # case "-.": option.RemoveKeywordsFromFullList(SubtractedKeywords.PickImplied)
        # case "-A": option.RemoveKeywordsFromFullList(SubtractedKeywords.All)
        # case "-B": option.RemoveKeywordsFromFullList(SubtractedKeywords.BoxAuto)
        # case "-M": option.RemoveKeywordsFromFullList(SubtractedKeywords.Multiple)
        # case "-C": option.RemoveKeywordsFromFullList(SubtractedKeywords.CrossingCPolygon)
        # case "-F": option.RemoveKeywordsFromFullList(SubtractedKeywords.Fence)
        # case "-L": option.RemoveKeywordsFromFullList(SubtractedKeywords.Last)
        # case "-P": option.RemoveKeywordsFromFullList(SubtractedKeywords.Previous)
        # case "-W": option.RemoveKeywordsFromFullList(SubtractedKeywords.WindowWPolygon)
        # case "-G": option.RemoveKeywordsFromFullList(SubtractedKeywords.Group)
        # case "-D": option.RemoveKeywordsFromFullList(SubtractedKeywords.AddRemove)


    if dxfcode_filter_list != []:
        # value = [TypedValue(System.Int32(0), "Circle")] # == AutoLisp (DxfCode . "Circle") 
        typevalue_list = []
        for [dxfcode, checkchar] in dxfcode_filter_list:
            typevalue_list.append(TypedValue(System.Int32(dxfcode), checkchar))
        filter = SelectionFilter(typevalue_list)
        result = ed.GetSelection(option, filter)
    else:
        result = ed.GetSelection(option)

    str(result)
    # result = ed.SelectImplied() #  need cad pickfirst set 1
    ss1 = result.Value # SelectionSet
    # ids = ss1.GetObjectIds()
    # ss2 = SelectionSet.FromObjectIds([ids[-1]])
    str(ss1) # 原理不明但有用，处理Error: eInvalidInput 在 Autodesk.AutoCAD.EditorInput.SelectionSetDelayMarshalled.GetObjectIds()
    if sssetfirst:
        ed.SetImpliedSelection(ss1) # Highlight(ss1) # (sssetfirst nil ss1)
    return ss1 

# 6. SelectImplied 选择当前图形中已经选择的实体
# 7. SelectLast 选择图形中最后一盒绘制的实体



def SSGetIdList(dxfcode_filter_list=[], sel_method="", string="", sssetfirst=False):
    ss1 = SSGet(dxfcode_filter_list, sel_method, string, sssetfirst)
    if ss1 == None: return []
    return [objid for objid in ss1.GetObjectIds()]

def SSGetDbObjectCollection(dxfcode_filter_list=[]):
    ss1 = SSGet(dxfcode_filter_list)
    if ss1 == None: return None
    with transaction() as trans:
        dbobj_collect = DBObjectCollection()
        for objid in ss1.GetObjectIds():
            objref_write = trans.GetObject(objid, OpenMode.ForWrite)
            dbobj_collect.Add(objref_write)

    return dbobj_collect

def SSSetFromIdList(objidlist=[]):
    if objidlist == []: return None
    ss2 = SelectionSet.FromObjectIds(objidlist)
    return ss2

def Highlight(ss1:SelectionSet):
    with transaction() as trans:
        for objid in ss1.GetObjectIds():
            objref = trans.GetObject(objid, OpenMode.ForRead)
            objref.Highlight()

def SSSetFirst(ss1):
    ed.SetImpliedSelection(ss1) # (sssetfirst nil ss1)


def GetExplode(objid):
    with transaction() as trans:
        entity = trans.GetObject(objid, OpenMode.ForWrite)
        result = DBObjectCollection()
        entity.Explode(result)
        for objref in result:
            AddDBObject(objref)
        entity.Erase()
    return [objref.ObjectId for objref in result]


def TransExplode(objid):
    entity = trans.GetObject(objid, OpenMode.ForWrite)
    result = DBObjectCollection()
    entity.Explode(result)
    for objref in result:
        AddDBObject(objref)
    entity.Erase()
    return [objref.ObjectId for objref in result]


# Point3d center = Point3d(1000, 0, 0)
# Vector3d normal = Vector3d(0, 0, 1)
# Circle circle1 = Circle(center, normal, 1000)
# circle1.ColorIndex = 1
# circle1.Thickness = 5
def TransCopy(objid, sourcept=[0,0,0], targetpt=[0,0,0], layer_name="", color_index=0):
    pt1, pt2 = Vec2toVec3(sourcept), Vec2toVec3(targetpt)
    dr1 = Direct(pt1, pt2)
    vecdr = Vector3d(*dr1)
    matrix4x4 = Matrix3d.Displacement(vecdr)
    entity = trans.GetObject(objid, OpenMode.ForRead)
    copyentity = entity.GetTransformedCopy(matrix4x4)
    AddDBObject(copyentity, layer_name, color_index)
    return copyentity


def TransRoationCopy(objid, angle, axis=[], center=[], layer_name="", color_index=0):
    center = ToPoint3d(center)
    axis = Vector3d(*axis)
    rad = angle * 0.01745329
    matrix4x4 = Matrix3d.Rotation(rad, axis, center)
    entity = trans.GetObject(objid, OpenMode.ForRead)
    copyentity = entity.GetTransformedCopy(matrix4x4)
    AddDBObject(copyentity, layer_name, color_index)
    return copyentity

def TransRoationCopyIdList(objidlist, angle, axis=[], center=[], layer_name="", color_index=0):
    result = []
    for objid in objidlist:
        copyref = TransRoationCopy(objid, angle, axis, center, layer_name, color_index)
        result.append(copyref.ObjectId)
    return result

def TransMoveIdList(objidlist, sourcept=[0,0,0], targetpt=[0,0,0], layer_name="", color_index=0):
    for objid in objidlist:
        TransMove(objid, sourcept, targetpt, layer_name, color_index)
    return objidlist


def TransMove(objid, sourcept=[0,0,0], targetpt=[0,0,0], layer_name="", color_index=0):
    pt1, pt2 = Vec2toVec3(sourcept), Vec2toVec3(targetpt)
    dr1 = Direct(pt1, pt2)
    vecdr = Vector3d(*dr1)
    matrix4x4 = Matrix3d.Displacement(vecdr)
    entity = trans.GetObject(objid, OpenMode.ForWrite)
    entity.TransformBy(matrix4x4)
    return entity


def TransRoation(objid, angle, axis=[], center=[], layer_name="", color_index=0):
    center = ToPoint3d(center)
    axis = Vector3d(*axis)
    rad = angle * 0.01745329
    matrix4x4 = Matrix3d.Rotation(rad, axis, center)
    entity = trans.GetObject(objid, OpenMode.ForWrite)
    entity.TransformBy(matrix4x4)
    return entity



def TransRoationIdList(objidlist, angle, axis=[], center=[], layer_name="", color_index=0):
    result = []
    for objid in objidlist:
        copyref = TransRoation(objid, angle, axis, center, layer_name, color_index)
        result.append(copyref.ObjectId)
    return result



def TransDimTextBoundXY0(objid):
    objref = trans.GetObject(objid, OpenMode.ForRead)
    block_table_record = trans.GetObject(objref.DimBlockId, OpenMode.ForRead)
    for objid in block_table_record: 
        objref = trans.GetObject(objid, OpenMode.ForRead)
        if type(objref) == MText:
            extend = objref.GeometricExtents
            point1 = extend.MinPoint
            point2 = extend.MaxPoint
            return [point1.X, point1.Y, 0], [point2.X, point2.Y, 0]

def TransDimLineObjectIdListBoundXY0(objidlist):
    bufidlist = []
    for objid in objidlist:
        objref = trans.GetObject(objid, OpenMode.ForRead)
        block_table_record = trans.GetObject(objref.DimBlockId, OpenMode.ForRead)
        for objid in block_table_record: 
            objref = trans.GetObject(objid, OpenMode.ForRead)
            if type(objref) == Line: bufidlist.append(objid)
    return TransObjectIdListBoundXY0(bufidlist)





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



def AddPoint(pt1, layer_name="", color_index=0):
    point = DBPoint(ToPoint3d(pt1))
    AddDBObject(point, layer_name, color_index)
    return point

def AddLine(start_point, final_point, layer_name="", color_index=0):
    line = Line(ToPoint3d(start_point), ToPoint3d(final_point))
    AddDBObject(line, layer_name, color_index)
    return line

def AddRect(ptmin, ptmax, layer_name="", color_index=0):
    x1, y1 = ptmin[0:2]
    x2, y2 = ptmax[0:2]
    pt1 = [x1, y1]
    pt2 = [x2, y1]
    pt3 = [x2, y2]
    pt4 = [x1, y2]
    pline = Polyline()
    for i, pt0 in enumerate([pt1, pt2, pt3, pt4]):
        pline.AddVertexAt(i, ToPoint2d(pt0), 0, 0, 0)
    pline.Closed = True
    AddDBObject(pline, layer_name, color_index)
    return pline


def AddLWPolyLine(ptlist, layer_name="", color_index=0):
    pline = Polyline()
    for i, pt1 in enumerate(ptlist):
        pline.AddVertexAt(i, ToPoint2d(pt1), 0, 0, 0)
    AddDBObject(pline, layer_name, color_index)
    return pline

def AddMKPolyLine(ptlist, layer_name="", color_index=0):
    pline = Polyline()
    for i, pt1 in enumerate(ptlist):
        pline.AddVertexAt(i, ToPoint2d(pt1), 0, 0, 0)
    pline.Closed = True
    AddDBObject(pline, layer_name, color_index)
    return pline


def AddPointCloud(pt1, color_rgb=[255,255,255]):
    r, g, b = color_rgb
    point = DBPoint(ToPoint3d(pt1))
    point.Color = Color.FromRgb(System.Byte(r), System.Byte(g), System.Byte(b))
    AddDBObject(point)
    return point



def AddText(pt1, string="单行文字", size=50, angle=0, layer_name="", color_index=0):
    text = DBText()
    text.Position = ToPoint3d(pt1) 
    text.TextString = string
    text.Height = size
    text.Rotation = Angle2Rad(angle)
    # text.IsMirroredInX = True # 在X轴镜像
    # text.HorizontalMode = TextHorizontalMode.TextCenter # 设置对齐方式
    # text.AlignmentPoint = text.Position # 设置对齐点
    AddDBObject(text, layer_name, color_index)
    return text




# Autodesk.AutoCAD.DatabaseServices.AlignedDimension
# DIMALIGNED
def AddDal(pt1, pt2, pt3, dimstylename="", layer_name="", color_index=0):
    return AddAlignedDimension(pt1, pt2, pt3, dimstylename, layer_name, color_index)
def AddAlignedDimension(pt1, pt2, pt3, dimstylename="", layer_name="", color_index=0):
    dal = AlignedDimension()
    dal.XLine1Point = ToPoint3d(pt1)
    dal.XLine2Point = ToPoint3d(pt2)
    dal.DimLinePoint = ToPoint3d(pt3)
    # dal.TextPosition = txtpt,
    dal.DimensionStyle = TransFindDimStyle(dimstylename) # 自带 if dimstylename != "": 判断
    AddDBObject(dal, layer_name, color_index)
    return dal

def AddDalLinear(pt1, pt2, direct="+x", dle=50, dimflagnum=50):
    x1, y1 = pt1[0], pt1[1]
    x2, y2 = pt2[0], pt2[1]
    try: z1 = pt1[2] 
    except: pass
    try: z2 = pt2[2] 
    except: pass
    xm, ym, zm = [(x1+x2)/2, (y1+y2)/2, (z1+z2)/2]
    mid = [xm, ym, zm]

    if dle < 0: raise ValueError("dle must >= 0 ... ... ")
    pt1 = [x1, y1, z1]
    pt2 = [x2, y2, z2]
    if direct == "+x" : pt3 = [xm+dle, ym, zm]
    if direct == "-x" : pt3 = [xm-dle, ym, zm]
    if direct == "+y" : pt3 = [xm, ym+dle, zm]
    if direct == "-y" : pt3 = [xm, ym-dle, zm]
    dr1 = GetPerDirectResetLengthXY(pt1, pt2, pt3, dle)
    pt3 = Vec3Add(mid, dr1)
    dimstylename = f"副本{dimflagnum} ISO-25"
    AddAlignedDimension(pt1, pt2, pt3, dimstylename)

# Autodesk.AutoCAD.DatabaseServices.RotatedDimension
# DIMLINEAR
def AddDim(pt1, pt2, pt3, dimstylename="", layer_name="", color_index=0):
    return AddRotatedDimension(pt1, pt2, pt3, dimstylename, layer_name, color_index)
def AddRotatedDimension(pt1, pt2, pt3, dimstylename="", layer_name="", color_index=0):
    angle = AngleFromCrossDr1Dr2(Direct(pt1, pt2), [1,0,0])
    if angle <  45: angle = 0
    if angle >= 45: angle = 90
    dim = RotatedDimension()
    dim.XLine1Point = ToPoint3d(pt1)
    dim.XLine2Point = ToPoint3d(pt2)
    dim.Rotation = Angle2Rad(angle)
    dim.DimLinePoint = ToPoint3d(pt3)
    # # dim.TextPosition = txtpt,
    # dim.DimensionText = txtpt,
    dim.DimensionStyle = TransFindDimStyle(dimstylename) # dim.DimensionStyle need ObjectId
    AddDBObject(dim, layer_name, color_index)
    return dim


def AddDimLinear(pt1, pt2, direct="+x", dle=50, dimflagnum=50):
    x1, y1 = pt1[0], pt1[1]
    x2, y2 = pt2[0], pt2[1]
    try: z1 = pt1[2] 
    except: pass
    try: z2 = pt2[2] 
    except: pass
    min_x, min_y, max_x, max_y = min(x1, x2), min(y1, y2), max(x1, x2), max(y1, y2)

    if dle < 0: raise ValueError("dle must >= 0 ... ... ")
    dimstylename = f"副本{dimflagnum} ISO-25"
    if direct == "+x" :
        pt3 = [max_x+dle, min_y]
        AddRotatedDimension(pt1, pt2, pt3, dimstylename)
    if direct == "-x" :
        pt3 = [min_x-dle, min_y]
        AddRotatedDimension(pt1, pt2, pt3, dimstylename)
    if direct == "+y" :
        pt3 = [min_x, max_y+dle]
        AddRotatedDimension(pt1, pt2, pt3, dimstylename)
    if direct == "-y" :
        pt3 = [max_x, min_y-dle]
        AddRotatedDimension(pt1, pt2, pt3, dimstylename)

    if direct != "+x" and direct != "-x" and direct != "+y" and direct != "-y": raise ValueError("DimLinear::direct::must be +x or -x or +y or -y , not {direct}")




def AddGroup(objidlist):
    groupdict = trans.GetObject(db.GroupDictionaryId, OpenMode.ForWrite)
    name = CalcGroupName()
    group = Group(name, True)
    for objid in objidlist:
        group.Append(objid)
    objid = groupdict.SetAt(name, group)
    trans.AddNewlyCreatedDBObject(group, True)
    return group


def AddPolyline3d(ptlist, layer_name="", color_index=0):
    collection = Point3dCollection()
    for pt1 in ptlist:
        collection.Add(ToPoint3d(pt1))
    pline = Polyline3d(Poly3dType.SimplePoly, collection, False) # Closed = True
    AddDBObject(pline, layer_name, color_index)
    return pline

def AddMKPolyline3d(ptlist, layer_name="", color_index=0):
    collection = Point3dCollection()
    for pt1 in ptlist:
        collection.Add(ToPoint3d(pt1))
    pline = Polyline3d(Poly3dType.SimplePoly, collection, True) # Closed = True
    AddDBObject(pline, layer_name, color_index)
    return pline



def AddRegion(dbobj_collect:DBObjectCollection, layer_name="", color_index=0):
    regions = Region.CreateFromCurves(dbobj_collect)
    for objref in regions:
        AddDBObject(objref, layer_name, color_index)
    return regions


def AddCircle(center, radius, normal=[0,0,1], layer_name="", color_index=0):
    circle = Circle()
    circle.Center = ToPoint3d(center)
    circle.Normal = ToVector3d(normal)
    circle.Radius = radius
    AddDBObject(circle, layer_name, color_index)
    return circle


def AddDBObject(dbobjref, layer_name="", color_index=0):
    CheckLayerAndColor(dbobjref, layer_name, color_index)
    objid = currentblock.AppendEntity(dbobjref)
    trans.AddNewlyCreatedDBObject(dbobjref, True)
    return objid


def DBObjectLine(start_point, final_point):
    line = Line(ToPoint3d(start_point), ToPoint3d(final_point))
    return line


def DBObjectRect(ptmin, ptmax):
    x1, y1 = ptmin[0:2]
    x2, y2 = ptmax[0:2]
    pt1 = [x1, y1]
    pt2 = [x2, y1]
    pt3 = [x2, y2]
    pt4 = [x1, y2]
    pline = Polyline()
    for i, pt0 in enumerate([pt1, pt2, pt3, pt4, pt1]):
        pline.AddVertexAt(i, ToPoint2d(pt0), 0, 0, 0)
    return pline


def DBObjectLWPolyLine(ptlist):
    pline = Polyline()
    for i, pt1 in enumerate(ptlist):
        pline.AddVertexAt(i, ToPoint2d(pt1), 0, 0, 0)
    return pline


def DBObjectMKPolyLine(ptlist, layer_name="", color_index=0):
    pline = Polyline()
    for i, pt1 in enumerate(ptlist):
        pline.AddVertexAt(i, ToPoint2d(pt1), 0, 0, 0)
    pline.Closed = True
    return pline


def DBObjectCopy(objref):
    matrix4x4 = Matrix3d.Identity # 单位矩阵
    copyobjref = objref.GetTransformedCopy(matrix4x4)
    return copyobjref




zhu_block_name_count = 1
def CalcBlockName():
    global zhu_block_name_count
    timestr = time.strftime("%Y%m%d%H%M%S")
    namestr = f"ZHUB"+timestr+f"{zhu_block_name_count:03d}"
    # Prompt(namestr)
    zhu_block_name_count += 1
    if zhu_block_name_count > 999: zhu_block_name_count = 1
    return namestr


zhu_group_name_count = 1
def CalcGroupName():
    global zhu_group_name_count
    timestr = time.strftime("%Y%m%d%H%M%S")
    namestr = f"ZHUG"+timestr+f"{zhu_group_name_count:03d}"
    # Prompt(namestr)
    zhu_group_name_count += 1
    if zhu_group_name_count > 999: zhu_group_name_count = 1
    return namestr


def AddBlockFromIdList(objidlist, base_point=[0,0,0]):
    # 创建块
    block = BlockTableRecord()
    block.Name = CalcBlockName()
    block.Origin = ToPoint3d(base_point)
    for objid in objidlist:
        objref = trans.GetObject(objid, OpenMode.ForWrite)
        block.AppendEntity(objref.Clone()), objref.Erase()
        # block.AppendEntity(objref) # eAlreadyInDb 在 Autodesk.AutoCAD.DatabaseServices.BlockTableRecord.AppendEntity(Entity entity)
    # 需要先往block里添加对象，才能往blocktable和db里添加对象，这样才能正确显示块
    blocktable = trans.GetObject(db.BlockTableId, OpenMode.ForWrite)
    blockobjid = blocktable.Add(block)
    trans.AddNewlyCreatedDBObject(block, True)
    # 插入块
    blockref = BlockReference(block.Origin, blockobjid)
    AddDBObject(blockref)
    return blockref

def AddBlockFromRefList(objreflist, base_point=[0,0,0]):
    # 创建块
    block = BlockTableRecord()
    block.Name = CalcBlockName()
    block.Origin = ToPoint3d(base_point)
    for objref in objreflist:
        block.AppendEntity(objref)
        # block.AppendEntity(objref) # eAlreadyInDb 在 Autodesk.AutoCAD.DatabaseServices.BlockTableRecord.AppendEntity(Entity entity)
    # 需要先往block里添加对象，才能往blocktable和db里添加对象，这样才能正确显示块
    blocktable = trans.GetObject(db.BlockTableId, OpenMode.ForWrite)
    blockobjid = blocktable.Add(block)
    trans.AddNewlyCreatedDBObject(block, True)
    # 插入块
    blockref = BlockReference(block.Origin, blockobjid)
    AddDBObject(blockref)
    return blockref


def TransObjectForRead(objid):
    objref = trans.GetObject(objid, OpenMode.ForRead)
    return objref


def TransObjectForWrite(objid):
    objref = trans.GetObject(objid, OpenMode.ForWrite)
    return objref


def TransFindDimStyle(stylename=""):
    if stylename == "": return db.Dimstyle
    dim_style_table = trans.GetObject(db.DimStyleTableId, OpenMode.ForRead)
    if dim_style_table.Has(stylename):
        dim_style_table_record_objid = dim_style_table[stylename]
        # dim_style_table_record = trans.GetObject(dim_style_table_record_objid, OpenMode.ForRead)
        return dim_style_table_record_objid
    else:
        return db.Dimstyle
# Autodesk.AutoCAD.DatabaseServices.DimStyleTableRecord value cannot be converted to Autodesk.AutoCAD.DatabaseServices.ObjectId

def TransCurrentDimStyle(stylename=""):
    dim_style_table = trans.GetObject(db.DimStyleTableId, OpenMode.ForRead)
    if dim_style_table.Has(stylename):
        dim_style_table_record_objid = dim_style_table[stylename]
        dim_style_table_record = trans.GetObject(dim_style_table_record_objid, OpenMode.ForRead)
        db.Dimstyle = dim_style_table_record_objid
        db.SetDimstyleData(dim_style_table_record) # 不设置这行，显示效果与标注名称应有的效果不一致
        objref = dim_style_table_record
    else:
        objref = trans.GetObject(db.Dimstyle, OpenMode.ForRead)
    return objref.ObjectId
    
def SetCurrentDimStyle(stylename=""):
    with transaction() as trans:
        dim_style_table = trans.GetObject(db.DimStyleTableId, OpenMode.ForRead)
        if dim_style_table.Has(stylename):
            dim_style_table_record_objid = dim_style_table[stylename]
            dim_style_table_record = trans.GetObject(dim_style_table_record_objid, OpenMode.ForRead)
            db.Dimstyle = dim_style_table_record_objid
            db.SetDimstyleData(dim_style_table_record) # 不设置这行，显示效果与标注名称应有的效果不一致
            objref = dim_style_table_record
        else:
            objref = trans.GetObject(db.Dimstyle, OpenMode.ForRead)
    return objref.ObjectId


def TransCurrentDimStyle(stylename=""):
    if stylename == "": return db.Dimstyle
    dim_style_table = trans.GetObject(db.DimStyleTableId, OpenMode.ForRead)
    if dim_style_table.Has(stylename):
        dim_style_table_record_objid = dim_style_table[stylename]
        dim_style_table_record = trans.GetObject(dim_style_table_record_objid, OpenMode.ForRead)
        db.Dimstyle = dim_style_table_record_objid
        return dim_style_table_record_objid
    else:
        return db.Dimstyle


def TransLWPolyLinePointList(objid:ObjectId, pt1=None):
    pline = trans.GetObject(objid, OpenMode.ForRead)
    po1 = [pline.StartPoint.X, pline.StartPoint.Y, pline.StartPoint.Z]
    po2 = [pline.EndPoint.X, pline.EndPoint.Y, pline.EndPoint.Z]
    result = []
    for i in range(pline.NumberOfVertices):
        point = pline.GetPoint3dAt(i)
        result.append([point.X, point.Y, point.Z])
    if pline.Closed: result = result + result[0:1]
    if pt1 == None: return result
    if Distance(pt1, po1) > Distance(pt1, po2): result = result[::-1]
    return result


def TransMKPolyLinePointList(objid:ObjectId):
    pline = trans.GetObject(objid, OpenMode.ForRead)
    result = []
    if pline.Closed:
        for i in range(pline.NumberOfVertices):
            point = pline.GetPoint3dAt(i)
            result.append([point.X, point.Y, point.Z])
    return result


def TransMKPolyLineDirectList(objid:ObjectId):
    ptlist = TransMKPolyLinePointList(objid)
    if ptlist == []: return []
    ptlist = ptlist + ptlist[0:1]
    drlist = []
    for i in range(len(ptlist)-1):
        pt1 = ptlist[i]
        pt2 = ptlist[i+1]
        dr1 = Direct(pt1, pt2)
        drlist.append(dr1)
    return drlist




def TransLWPolyLineDirectList(objid:ObjectId, pt1=None):
    ptlist = TransLWPolyLinePointList(objid, pt1)
    drlist = []
    for i in range(len(ptlist)-1):
        pt1 = ptlist[i]
        pt2 = ptlist[i+1]
        dr1 = Direct(pt1, pt2)
        drlist.append(dr1)
    return drlist

def TransLWPolyLineDotAngleList(objid:ObjectId, pt1=None):
    ptlist = TransLWPolyLinePointList(objid, pt1)
    anglelist = [None]
    for i in range(1, len(ptlist)-1):
        pt1 = ptlist[i-1]
        pt2 = ptlist[i]
        pt3 = ptlist[i+1]
        dr1 = Direct(pt2, pt1)
        dr2 = Direct(pt2, pt3)
        angle = AngleFromDotDr1Dr2(dr1, dr2)
        anglelist.append(angle)
    anglelist.append(None)
    return anglelist



def TransLWPolyLineLengthList(objid:ObjectId, pt1=None):
    ptlist = TransLWPolyLinePointList(objid, pt1)
    lnlist = []
    for i in range(len(ptlist)-1):
        pt1 = ptlist[i]
        pt2 = ptlist[i+1]
        ln1 = Distance(pt1, pt2)
        lnlist.append(ln1)
    return lnlist

def TransNRectPointListFromSheet(objid1:ObjectId, objid2:ObjectId, pt1=None):
    if pt1 == None:
        pt1, pt2 = GetIdListBoundXY0([objid1, objid2])
        x1, y1, z1 = pt1
        x2, y2, z2 = pt2
        pt1 = [x1, y2, 0]
    edgelist1 = TransLWPolyLineEdgeList(objid1, pt1)
    edgelist2 = TransLWPolyLineEdgeList(objid2, pt1)
    # ptlist1 = TransLWPolyLinePointList(objid1, pt1)
    # ptlist2 = TransLWPolyLinePointList(objid1, pt1)
    # pt1, pt2 = ptlist1[0], ptlist1[-1]
    # po1, po2 = ptlist2[0], ptlist2[-1]
    # if Distance(pt1, po1) > Distance(pt1, po2): ptlist2 = ptlist2[::-1]
    # edgelist1 = []
    # for i in range(len(ptlist1)-1):
    #     pt1 = ptlist1[i]
    #     pt2 = ptlist1[i+1]
    #     edgelist1.append([pt1, pt2])
    # edgelist2 = []
    # for i in range(len(ptlist2)-1):
    #     po1 = ptlist2[i]
    #     po2 = ptlist2[i+1]
    #     edgelist2.append([po1, po2])

    buflist = []
    for [pt1, pt2], [po1, po2] in zip(edgelist1, edgelist2):
        # 平行线1
        line1 = Line(ToPoint3d(pt1), ToPoint3d(pt2))
        point = line1.GetClosestPointTo(ToPoint3d(po1), extend=False) # 会返回直线端点 # System.Boolean.Parse("False") 
        pd1 = [point.X, point.Y, point.Z]
        point = line1.GetClosestPointTo(ToPoint3d(po2), extend=False) 
        pd2 = [point.X, point.Y, point.Z]
        # 平行线2
        line2 = Line(ToPoint3d(po1), ToPoint3d(po2))
        point = line2.GetClosestPointTo(ToPoint3d(pt1), extend=False) # 会返回直线端点 # System.Boolean.Parse("False") 
        pn1 = [point.X, point.Y, point.Z]
        point = line2.GetClosestPointTo(ToPoint3d(pt2), extend=False) 
        pn2 = [point.X, point.Y, point.Z]
        buflist.append([pd1, pd2, pn2, pn1])
    return buflist

def TransNRectLengthListFromSheet(objid1:ObjectId, objid2:ObjectId, pt1=None):
    buflist = TransNRectPointListFromSheet(objid1, objid2, pt1)
    lengthlist, nesslist = [], []
    for pt1, pt2, pt3, pt4 in buflist:
        length = Distance(pt1, pt2)
        ness = Distance(pt2, pt3)
        lengthlist.append(length)
        nesslist.append(ness)   
    return lengthlist, ness

def ShapLengthFromDeepthAndAngle(depth, angle):
    angle = angle / 2 # tan = sin / cos
    length = depth / math.tan(Angle2Rad(angle))
    return length

def ShapLengthFromDeepthAndPt1Pt2Pt3(depth, pt1, pt2, pt3):
    dr1 = Direct(pt2, pt1)
    dr2 = Direct(pt2, pt3)
    angle = AngleFromDotDr1Dr2(dr1, dr2)
    angle = angle / 2 # tan = sin / cos
    length = depth / math.tan(Angle2Rad(angle))
    return length


def TransLWPolyLineEdgeList(objid:ObjectId, pt1=None):
    ptlist = TransLWPolyLinePointList(objid, pt1)
    buflist = []
    for i in range(len(ptlist)-1):
        pt1 = ptlist[i]
        pt2 = ptlist[i+1]
        buflist.append([pt1, pt2])
    return buflist



# def TransLWPolyLineArea(objid:ObjectId):
#     dbobj_collect = DBObjectCollection()
#     matrix4x4 = Matrix3d.Identity # 单位矩阵
#     objref = trans.GetObject(objid, OpenMode.ForRead)
#     copyobjref = objref.GetTransformedCopy(matrix4x4)
#     dbobj_collect.Add(copyobjref)
#     regions = Region.CreateFromCurves(dbobj_collect)
#     region = regions[0]
#     return region.Area

def TransEntityStartMidEndPoint(objid:ObjectId):
    objref = trans.GetObject(objid, OpenMode.ForRead)
    point1 = objref.StartPoint
    point2 = objref.EndPoint
    pt1 = [point1.X, point1.Y, point1.Z]
    mid = [(point1.X+point2.X)/2, (point1.Y+point2.Y)/2, (point1.Z+point2.Z)/2]
    pt2 = [point2.X, point2.Y, point2.Z]
    return pt1, mid, pt2

def TransEntityStartEndPoint(objid:ObjectId):
    objref = trans.GetObject(objid, OpenMode.ForRead)
    point1 = objref.StartPoint
    point2 = objref.EndPoint
    return [point1.X, point1.Y, point1.Z], [point2.X, point2.Y, point2.Z]

def TransStartPoint(objid):
    objref = trans.GetObject(objid, OpenMode.ForRead)
    start = objref.StartPoint
    return [start.X, start.Y, start.Z]


def TransStartEndPoint(objid):
    objref = trans.GetObject(objid, OpenMode.ForRead)
    start = objref.StartPoint
    final = objref.EndPoint
    # Prompt([start, final])
    return [start.X, start.Y, start.Z], [final.X, final.Y, final.Z]


def TransLWPolyLineStartMid(objid:ObjectId):
    pline = trans.GetObject(objid, OpenMode.ForRead)
    point1 = pline.GetPoint3dAt(0)
    point2 = pline.GetPoint3dAt(1)
    pt1 = [point1.X, point1.Y, point1.Z]
    pt2 = [point2.X, point2.Y, point2.Z]
    mid = MidPt1Pt2(pt1, pt2)
    return mid


def TransLWPolyLineNormal(objid:ObjectId):
    pline = trans.GetObject(objid, OpenMode.ForRead)
    return pline.Normal


def TransEntityBoundXYZ(objid:ObjectId):
    objref = trans.GetObject(objid, OpenMode.ForRead)
    extend = objref.GeometricExtents
    point1 = extend.MinPoint
    point2 = extend.MaxPoint
    return [point1.X, point1.Y, point1.Z], [point2.X, point2.Y, point2.Z]


def TransEntityBoundCenterXYZ(objid:ObjectId):
    objref = trans.GetObject(objid, OpenMode.ForRead)
    extend = objref.GeometricExtents
    point1 = extend.MinPoint
    point2 = extend.MaxPoint
    return [(point1.X+point2.X)/2, (point1.Y+point2.Y)/2, (point1.Z+point2.Z)/2]


def TransEntityBoundXY0(objid:ObjectId):
    objref = trans.GetObject(objid, OpenMode.ForRead)
    extend = objref.GeometricExtents
    point1 = extend.MinPoint
    point2 = extend.MaxPoint
    return [point1.X, point1.Y, 0], [point2.X, point2.Y, 0]


def TransEntityBoundCenterXY0(objid:ObjectId):
    objref = trans.GetObject(objid, OpenMode.ForRead)
    extend = objref.GeometricExtents
    point1 = extend.MinPoint
    point2 = extend.MaxPoint
    return [(point1.X+point2.X)/2, (point1.Y+point2.Y)/2, 0]


def TransEntityBoundLengthWidth(objid:ObjectId):
    objref = trans.GetObject(objid, OpenMode.ForRead)
    extend = objref.GeometricExtents
    point1 = extend.MinPoint
    point2 = extend.MaxPoint
    return point2.X-point1.X, point2.Y-point1.Y


def TransEntityBoundLengthWidthHeight(objid:ObjectId):
    objref = trans.GetObject(objid, OpenMode.ForRead)
    extend = objref.GeometricExtents
    point1 = extend.MinPoint
    point2 = extend.MaxPoint
    return point2.X-point1.X, point2.Y-point1.Y, point2.Z-point1.Z


def TransEntityArea(objid:ObjectId):
    objref = trans.GetObject(objid, OpenMode.ForRead)
    area = objref.Area
    return area






def TransEntityLength(objid:ObjectId):
    objref = trans.GetObject(objid, OpenMode.ForRead)
    if type(objref) == Spline:
        para = objref.EndParam
        length = objref.GetDistanceAtParameter(para)
    else:
        length = objref.Length
    return length


def TransErase(objid:ObjectId):
    objref = trans.GetObject(objid, OpenMode.ForWrite)
    objref.Erase()


def TransSSBoundXYZ(ss1:SelectionSet):
    extend = Extents3d()
    for objid in ss1.GetObjectIds():
        objref = trans.GetObject(objid, OpenMode.ForRead)
        extend.AddExtents(objref.GeometricExtents) 
    point1 = extend.MinPoint
    point2 = extend.MaxPoint
    return [point1.X, point1.Y, point1.Z], [point2.X, point2.Y, point2.Z]


def TransSSBoundCenterXYZ(ss1:SelectionSet):
    extend = Extents3d()
    for objid in ss1.GetObjectIds():
        objref = trans.GetObject(objid, OpenMode.ForRead)
        extend.AddExtents(objref.GeometricExtents) 
    point1 = extend.MinPoint
    point2 = extend.MaxPoint
    return [(point1.X+point2.X)/2, (point1.Y+point2.Y)/2, (point1.Z+point2.Z)/2]


def TransSSBoundXY0(ss1:SelectionSet):
    extend = Extents3d()
    for objid in ss1.GetObjectIds():
        objref = trans.GetObject(objid, OpenMode.ForRead)
        extend.AddExtents(objref.GeometricExtents) 
    point1 = extend.MinPoint
    point2 = extend.MaxPoint
    return [point1.X, point1.Y, 0], [point2.X, point2.Y, 0]


def TransSSBoundCenterXY0(ss1:SelectionSet):
    extend = Extents3d()
    for objid in ss1.GetObjectIds():
        objref = trans.GetObject(objid, OpenMode.ForRead)
        extend.AddExtents(objref.GeometricExtents) 
    point1 = extend.MinPoint
    point2 = extend.MaxPoint
    return [(point1.X+point2.X)/2, (point1.Y+point2.Y)/2, 0]


def TransSSBoundLengthWidth(ss1:SelectionSet):
    extend = Extents3d()
    for objid in ss1.GetObjectIds():
        objref = trans.GetObject(objid, OpenMode.ForRead)
        extend.AddExtents(objref.GeometricExtents) 
    point1 = extend.MinPoint
    point2 = extend.MaxPoint
    return point2.X-point1.X, point2.Y-point1.Y


def TransSSBoundLengthWidthHeight(ss1:SelectionSet):
    extend = Extents3d()
    for objid in ss1.GetObjectIds():
        objref = trans.GetObject(objid, OpenMode.ForRead)
        extend.AddExtents(objref.GeometricExtents) 
    point1 = extend.MinPoint
    point2 = extend.MaxPoint
    return point2.X-point1.X, point2.Y-point1.Y, point2.Z-point1.Z


def TransObjectIdListBoundXYZ(objidlist):
    extend = Extents3d()
    for objid in objidlist:
        objref = trans.GetObject(objid, OpenMode.ForRead)
        extend.AddExtents(objref.GeometricExtents) 
    point1 = extend.MinPoint
    point2 = extend.MaxPoint
    return [point1.X, point1.Y, point1.Z], [point2.X, point2.Y, point2.Z]


def TransObjectIdListBoundCenterXYZ(objidlist):
    extend = Extents3d()
    for objid in objidlist:
        objref = trans.GetObject(objid, OpenMode.ForRead)
        extend.AddExtents(objref.GeometricExtents) 
    point1 = extend.MinPoint
    point2 = extend.MaxPoint
    return [(point1.X+point2.X)/2, (point1.Y+point2.Y)/2, (point1.Z+point2.Z)/2]


def TransObjectIdListBoundXY0(objidlist):
    extend = Extents3d()
    for objid in objidlist:
        objref = trans.GetObject(objid, OpenMode.ForRead)
        extend.AddExtents(objref.GeometricExtents) 
    point1 = extend.MinPoint
    point2 = extend.MaxPoint
    return [point1.X, point1.Y, 0], [point2.X, point2.Y, 0]


def TransObjectIdListBoundCenterXY0(objidlist):
    extend = Extents3d()
    for objid in objidlist:
        objref = trans.GetObject(objid, OpenMode.ForRead)
        extend.AddExtents(objref.GeometricExtents) 
    point1 = extend.MinPoint
    point2 = extend.MaxPoint
    return [(point1.X+point2.X)/2, (point1.Y+point2.Y)/2, 0]


def TransObjectIdListBoundLengthWidth(objidlist):
    extend = Extents3d()
    for objid in objidlist:
        objref = trans.GetObject(objid, OpenMode.ForRead)
        extend.AddExtents(objref.GeometricExtents) 
    point1 = extend.MinPoint
    point2 = extend.MaxPoint
    return point2.X-point1.X, point2.Y-point1.Y


def TransObjectIdListBoundLengthWidthHeight(objidlist):
    extend = Extents3d()
    for objid in objidlist:
        objref = trans.GetObject(objid, OpenMode.ForRead)
        extend.AddExtents(objref.GeometricExtents) 
    point1 = extend.MinPoint
    point2 = extend.MaxPoint
    return point2.X-point1.X, point2.Y-point1.Y, point2.Z-point1.Z


def BoundPointToRectPointListXY0(ptmin, ptmax):
    x1, y1 = ptmin[0:2]
    x2, y2 = ptmax[0:2]
    pt1 = [x1, y1, 0]
    pt2 = [x2, y1, 0]
    pt3 = [x2, y2, 0]
    pt4 = [x1, y2, 0]
    return pt1, pt2, pt3, pt4

def BoundLengthWidthToRectPointListXY0(pt1, length, width):
    # 逆时针
    x1, y1 = pt1[0:2]
    pt1 = [x1,        y1, 0]
    pt2 = [x1+length, y1, 0]
    pt3 = [x1+length, y1+width, 0]
    pt4 = [x1,        y1+width, 0]
    return pt1, pt2, pt3, pt4


def BoundLengthWidthToOffsetRectPointListXY0(pt1, length, width, offset):
    # 逆时针
    x1, y1 = pt1[0:2]
    pt1 = [-offset+x1,        -offset+y1, 0]
    pt2 = [+offset+x1+length, -offset+y1, 0]
    pt3 = [+offset+x1+length, +offset+y1+width, 0]
    pt4 = [-offset+x1,        +offset+y1+width, 0]
    return pt1, pt2, pt3, pt4




def IsPointInRange(pt1, objid:ObjectId|SelectionSet|Region):
    regions = []
    objtype = str(objid.GetType())
    match objtype:
        case "Autodesk.AutoCAD.DatabaseServices.ObjectId":
            with transaction() as trans:
                objref = trans.GetObject(objid, OpenMode.ForRead)
            dbobj_collect = DBObjectCollection()
            dbobj_collect.Add(objref)
            regions = Region.CreateFromCurves(dbobj_collect)
        case _: pass

    flag = False
    for region in regions:
        brep = Brep(region)
        result = brep.GetPointContainment(ToPoint3d(pt1), PointContainment.Outside) # out ref 被自动处理成结果返回
        # print(brep_entity.GetType())
        # (<Autodesk.AutoCAD.BoundaryRepresentation.Face object at 0x00000204409347C0>, <PointContainment.OnBoundary: 2>)
        # (None, <PointContainment.Outside: 1>)
        if result[1] == PointContainment.OnBoundary: 
            flag = True
            break   
    return flag



def IsInclude(objid1, objid2): # 交集
    with transaction() as trans:
        objref1 = trans.GetObject(objid1, OpenMode.ForRead)
        objref2 = trans.GetObject(objid2, OpenMode.ForRead)
    dbobj_collect1 = DBObjectCollection()
    dbobj_collect1.Add(objref1)
    dbobj_collect2 = DBObjectCollection()
    dbobj_collect2.Add(objref2)
    region1 = Region.CreateFromCurves(dbobj_collect1)[0]
    region2 = Region.CreateFromCurves(dbobj_collect2)[0]
    area1 = region1.Area
    area2 = region2.Area
    regionmax, regionmin = region1, region2
    areamax, areamin = area1, area2
    if area1 < area2: 
        regionmax, regionmin = region2, region1
        areamax, areamin = area2, area1
    regionmax.BooleanOperation(BooleanOperationType.BoolIntersect, regionmin) # Void函数返回None
    area3 = regionmax.Area # AutoCAD规定，布尔后的对象ID与原先面积大的一致

NoneType = type(None)
def IsNoneObjectId(objid):
    return IsNone(objid)
def IsNone(objid):
    if type(objid) == NoneType: return True
    if type(objid) == ObjectId:
        if objid.IsNull: return True
    return False




def IsPointInRect(pt0, ptlist=[]):
    pt1, pt2, pt3, pt4 = ptlist
    flag1 = GetPerflagXY(pt1, pt2, pt0)
    flag2 = GetPerflagXY(pt2, pt3, pt0)
    flag3 = GetPerflagXY(pt3, pt4, pt0)
    flag4 = GetPerflagXY(pt4, pt1, pt0)
    if flag1 >= 0 and flag2 >= 0 and flag3 >= 0 and flag4 >= 0 : return True
    if flag1 <= 0 and flag2 <= 0 and flag3 <= 0 and flag4 <= 0 : return True
    return False



def IsRectInRect(ptnlist=[], ptwlist=[]):
    pn1, pn2, pn3, pn4 = ptnlist
    # pw1, pw2, pw3, pw4 = pt2list
    flag1 = IsPointInRect(pn1, ptwlist)
    flag2 = IsPointInRect(pn2, ptwlist)
    flag3 = IsPointInRect(pn3, ptwlist)
    flag4 = IsPointInRect(pn4, ptwlist)
    if flag1 and flag2 and flag3 and flag4: return True
    return False


def IsCCW(ptlist=[]): # CounterClockWise 逆时针
    # 而对于一般的简单多边形，则需对于多边形的每一个点计算Cross值，如果正值比较多，是逆时针；负值较多则为顺时针。
    count = len(ptlist)
    if count <= 2: raise ValueError("点的数量少于3个")
    flagsum = 0
    for i in range(1, count-1):
        pt1 = ptlist[i-1]
        pt2 = ptlist[i]
        pt3 = ptlist[i+1]
        flagper = GetPerflagXY(pt1, pt2, pt3)
        flagsum += flagper
    if flagsum > 0: return True
    if flagsum < 0: return False
    raise ValueError("当前算法无法判断点集方向是顺时针还是逆时针")


def IsPointIsSolidVertex(pt0, objid):
    try: x0, y0, z0 = pt0
    except: [x0, y0], z0 = pt0, 0
    with transaction() as trans:
        objref = trans.GetObject(objid, OpenMode.ForWrite)
    brep = Brep(objref)
    flag0 = False
    for vertex in brep.Vertices:
        x1, y1, z1 = vertex.Point.X, vertex.Point.Y, vertex.Point.Z 
        flag1, flag2, flag3 = False, False, False
        if abs(x0-x1) < 0.00001: flag1 = True
        if abs(y0-y1) < 0.00001: flag2 = True
        if abs(z0-z1) < 0.00001: flag3 = True
        if flag1 and flag2 and flag3: 
            flag0 = True
            break
    return flag0


def IsPointIsPointListPoint(pt0, ptlist):
    flag0 = False
    for pt1 in ptlist:
        if IsPointSame(pt0, pt1):
            flag0 = True
            break
    return flag0



def IsPointSame(pt1, pt2, precision=0.00001):
    try: x1, y1, z1 = pt1
    except: [x1, y1], z1 = pt1, 0
    try: x2, y2, z2 = pt2
    except: [x2, y2], z2 = pt2, 0
    flag1, flag2, flag3 = False, False, False
    if abs(x2-x1) < precision: flag1 = True
    if abs(y2-y1) < precision: flag2 = True
    if abs(z2-z1) < precision: flag3 = True
    if flag1 and flag2 and flag3: return True
    return False

def IsPointInLWPolyLinePointList(pt1, ptlist): # ptlist 来自GetLWPoly not GetMkPoly
    x, y = pt1[0:2]
    inside = False
    p1x, p1y = ptlist[0][0:2]
    for pt2 in ptlist:
        p2x, p2y = pt2[0:2]
        if y > min(p1y, p2y):
            if y <= max(p1y, p2y):
                if x <= max(p1x, p2x):
                    if abs(p1y-p2y) > 0.00001: # p1y != p2y
                        xinters = (y - p1y) * (p2x - p1x) / (p2y - p1y) + p1x
                    if abs(p1x-p2x) < 0.00001 or x <= xinters: # p1y == p2y
                        inside = not inside
        p1x, p1y = p2x, p2y
    return inside

def ToPoint2d(pt0):
    x, y = pt0[0:2]
    return Point2d(x, y)


def ToPoint3d(pt0):
    try: x, y, z = pt0
    except: [x, y], z = pt0, 0
    return Point3d(x, y, z)


# def ToVector2d(pt0):
#     x, y = pt0[0:2]
#     return Vector2d(x, y)


def ToVector3d(pt0):
    try: x, y, z = pt0
    except: [x, y], z = pt0, 0
    return Vector3d(x, y, z)


def Normalized(x, y, z = 0):
    a = x**2 + y**2 + z**2
    distance = math.sqrt(a)
    xn, yn, zn = x/distance, y/distance, z/distance
    return [xn, yn, zn]

def Vec3ResetLength(dr1, length):
    x, y, z = Normalized(*dr1)
    x, y, z = x*length, y*length, z*length
    return [x, y, z]

# def VecXcomponent(pt1):
#     pass
# def VecYcomponent(pt1):
#     pass
# def VecZcomponent(pt1):
#     pass

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


def Dot(dr1, dr2):
    x1, y1, z1 = Vec2toVec3(dr1)
    x2, y2, z2 = Vec2toVec3(dr2)
    return x1*x2 + y1*y2 + z1*z2

def DotNormalized(dr1, dr2):
    # a⋅b=∣a∣∗∣b∣∗cosθ
    a = Vec2toVec3(dr1)
    b = Vec2toVec3(dr2)
    x1, y1, z1 = Normalized(*a)
    x2, y2, z2 = Normalized(*b)
    return x1*x2 + y1*y2 + z1*z2

def Cross(dr1, dr2):
    # ∣a×b∣=∣a∣∗∣b∣∗sinθ
    x1, y1, z1 = Vec2toVec3(dr1)
    x2, y2, z2 = Vec2toVec3(dr2)
    x3 =   y1*z2 - y2*z1
    y3 = -(x1*z2 - x2*z1)
    z3 =   x1*y2 - x2*y1
    return [x3, y3, z3]

def CrossNormalized(dr1, dr2):
    # ∣a×b∣=∣a∣∗∣b∣∗sinθ
    a = Vec2toVec3(dr1)
    b = Vec2toVec3(dr2)
    x1, y1, z1 = Normalized(*a)
    x2, y2, z2 = Normalized(*b)
    x3 =   y1*z2 - y2*z1
    y3 = -(x1*z2 - x2*z1)
    z3 =   x1*y2 - x2*y1
    return [x3, y3, z3]

def Rad2Angle(rad):
    angle = rad * 57.2957795
    return angle

def Angle2Rad(angle):
    rad = angle * 0.01745329
    return rad


def DirectListToPointList(pt, drlist):
    pt1 = pt
    result = [pt1]
    for dr1 in drlist:
        pt1 = Vec3Add(pt1, dr1)
        result.append(pt1)
    return result



def AngleFromDotDr1Dr2(dr1=[1,0,0], dr2=[0,0,1]):
    cos = DotNormalized(dr1, dr2)
    rad = math.acos(cos)
    angle = rad * 57.2957795
    return angle



def AngleFromCrossDr1Dr2(dr1=[1,0,0], dr2=[0,0,1]):
    normal = CrossNormalized(dr1, dr2)
    length = Distance([0,0,0], normal)
    sin = length
    rad = math.asin(sin)
    angle = rad * 57.2957795
    return angle



def MatrixRotationPointList(angle, axis, center, ptlist=[]):
    # center = ToPoint3d(center)
    # axis = Vector3d(*axis) # direct
    # rad = angle * 0.01745329
    # matrix4x4 = Matrix3d.Rotation(rad, axis, center)
    # point = ToPoint3d(pt1)
    # print(point)
    # point1 = point.TransformBy(matrix4x4)
    # print(point)
    # point3 = point.ScaleBy(double scaleFactor, Point3d centerPoint)
    resultlist = []
    for pt1 in ptlist:
        point1 = ToPoint3d(pt1)
        point2 = point1.RotateBy(angle*0.01745329, Vector3d(*axis), ToPoint3d(center))
        pt1 = [point2.X, point2.Y, point2.Z]
        resultlist.append(pt1)
    return resultlist
        
        
def PointToPointNormalPlane(pt1, pt0, normal):
    # 平面 点法式
    # Ax + By + Cz - (Ax0 + By0 + Cz0) = 0
    x0, y0, z0 = pt0
    A , B , C  = normal
    D = -(A*x0 + B*y0 + C*z0)
    x1, y1, z1 = pt1
    t = (A*x1 + B*y1 + C*z1 + D)/(A**2 + B**2 + C**2)
    x = x1 - A*t
    y = y1 - B*t
    z = z1 - C*t
    return [x, y, z]


def GetAttachGapAGapBPt1Pt2(pt1, pt2, gapa, gapb):
    return GetAttachGapAGapBPointList(pt1, pt2, gapa, gapb)
def GetAttachGapAGapBPointList(pt1, pt2, gapa, gapb):
    dr1 = Direct(pt1, pt2)
    dr1 = Vec3ResetLength(dr1, gapa)
    dr2 = Direct(pt2, pt1)
    dr2 = Vec3ResetLength(dr2, gapb)    
    return [Vec3Add(pt1, dr1), Vec3Add(pt2, dr2)]


def GetAttachExtandAExtandBPt1Pt2(pt1, pt2, extenda, extendb):
    return GetAttachExtandAExtandBPointList(pt1, pt2, extenda, extendb)
def GetAttachExtandAExtandBPointList(pt1, pt2, extenda, extendb):
    dr1 = Direct(pt2, pt1)
    dr1 = Vec3ResetLength(dr1, extenda)
    dr2 = Direct(pt1, pt2)
    dr2 = Vec3ResetLength(dr2, extendb)    
    return [Vec3Add(pt1, dr1), Vec3Add(pt2, dr2)]



def GetAttachNDirectPt1Pt2(pt1, pt2, length):
    return GetAttachNDirectPointList(pt1, pt2, length)
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


def GetAttachWDirectPt1Pt2(pt1, pt2, length):
    return GetAttachWDirectPointList(pt1, pt2, length)
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
    if flag > 0: return  1
    if flag < 0: return -1
    return 0


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


def DirectToPerDirectXY(dr0, perflag):
    x, y, z = dr0
    match perflag:
        case  1: return [-y,  x,  z]
        case -1: return [ y, -x,  z]
        case  _: raise ValueError("...点在线上...")

def GetPerDirect2XY(pt1, pt2):
    direct  = Direct(pt1, pt2)
    perdr1 = DirectToPerDirectXY(direct,  1)
    perdr2 = DirectToPerDirectXY(direct, -1)
    return perdr1, perdr2

def GetPerDirect2ResetLengthXY(pt1, pt2, length):
    direct  = Direct(pt1, pt2)
    perdr1 = DirectToPerDirectXY(direct,  1)
    perdr2 = DirectToPerDirectXY(direct, -1)
    x1, y1, z1 = Normalized(*perdr1)
    x2, y2, z2 = Normalized(*perdr2)
    return [x1*length, y1*length, z1*length], [x2*length, y2*length, z2*length]




def GetStartFinalPoint(objid):
    with transaction() as trans:
        objref = trans.GetObject(objid, OpenMode.ForRead)
        start = objref.StartPoint
        final = objref.EndPoint
    return [start.X, start.Y, start.Z], [final.X, final.Y, final.Z]

def GetStartPoint(objid):
    with transaction() as trans:
        objref = trans.GetObject(objid, OpenMode.ForRead)
        start = objref.StartPoint
    return [start.X, start.Y, start.Z]

def GetFinalPoint(objid):
    with transaction() as trans:
        objref = trans.GetObject(objid, OpenMode.ForRead)
        final = objref.EndPoint
    return [final.X, final.Y, final.Z]


def GetEntityColorIndex(objid:ObjectId):
    with transaction() as trans:
        objref = trans.GetObject(objid, OpenMode.ForRead)
        color_index = objref.ColorIndex
    return color_index

def GetEntityLayerName(objid:ObjectId):
    with transaction() as trans:
        objref = trans.GetObject(objid, OpenMode.ForRead)
        layer_name = objref.Layer
    return layer_name


def GetIdListColorIndexList(objidlist=[]):
    with transaction() as trans:
        color_index_list = []
        for objid in objidlist:
            objref = trans.GetObject(objid, OpenMode.ForRead)
            color_index = objref.ColorIndex
            if color_index == 0: pass
            if color_index == 256:
                layer_name = objref.Layer
                layer_table = trans.GetObject(db.LayerTableId, OpenMode.ForRead)
                layer_record_id = layer_table[layer_name]
                layer_record = trans.GetObject(layer_record_id, OpenMode.ForRead)
                color_index = layer_record.Color.ColorIndex
            color_index_list.append(color_index)
    return color_index_list

def GetIdListLayerNameList(objidlist=[]):
    with transaction() as trans:
        layer_name_list = []
        for objid in objidlist:
            objref = trans.GetObject(objid, OpenMode.ForRead)
            layer_name = objref.Layer
            layer_name_list.append(layer_name)
    return layer_name_list


def GetEntityBoundXYZ(objid:ObjectId):
    with transaction() as trans:
        objref = trans.GetObject(objid, OpenMode.ForRead)
        extend = objref.GeometricExtents
        point1 = extend.MinPoint
        point2 = extend.MaxPoint
    return [point1.X, point1.Y, point1.Z], [point2.X, point2.Y, point2.Z]


def GetEntityBoundCenterXYZ(objid:ObjectId):
    with transaction() as trans:
        objref = trans.GetObject(objid, OpenMode.ForRead)
        extend = objref.GeometricExtents
        point1 = extend.MinPoint
        point2 = extend.MaxPoint
    return [(point1.X+point2.X)/2, (point1.Y+point2.Y)/2, (point1.Z+point2.Z)/2]


def GetEntityBoundXY0(objid:ObjectId):
    with transaction() as trans:
        objref = trans.GetObject(objid, OpenMode.ForRead)
        extend = objref.GeometricExtents
        point1 = extend.MinPoint
        point2 = extend.MaxPoint
    return [point1.X, point1.Y, 0], [point2.X, point2.Y, 0]

def GetEntityBoundCenterXY0(objid:ObjectId):
    with transaction() as trans:
        objref = trans.GetObject(objid, OpenMode.ForRead)
        extend = objref.GeometricExtents
        point1 = extend.MinPoint
        point2 = extend.MaxPoint
    return [(point1.X+point2.X)/2, (point1.Y+point2.Y)/2, 0]


def GetEntityBoundLengthWidth(objid:ObjectId):
    with transaction() as trans:
        objref = trans.GetObject(objid, OpenMode.ForRead)
        extend = objref.GeometricExtents
        point1 = extend.MinPoint
        point2 = extend.MaxPoint
    return point2.X-point1.X, point2.Y-point1.Y

def GetEntityBoundLengthWidthHeight(objid:ObjectId):
    with transaction() as trans:
        objref = trans.GetObject(objid, OpenMode.ForRead)
        extend = objref.GeometricExtents
        point1 = extend.MinPoint
        point2 = extend.MaxPoint
    return point2.X-point1.X, point2.Y-point1.Y, point2.Z-point1.Z

def GetEntityStartEndPoint(objid:ObjectId):
    with transaction() as trans:
        objref = trans.GetObject(objid, OpenMode.ForRead)
        point1 = objref.StartPoint
        point2 = objref.EndPoint
    return [point1.X, point1.Y, point1.Z], [point2.X, point2.Y, point2.Z]


def GetEntityArea(objid:ObjectId):
    with transaction() as trans:
        objref = trans.GetObject(objid, OpenMode.ForRead)
        area = objref.Area
    return area


def GetEntityLength(objid:ObjectId):
    with transaction() as trans:
        objref = trans.GetObject(objid, OpenMode.ForRead)
        if type(objref) == Spline:
            para = objref.EndParam
            length = objref.GetDistanceAtParameter(para)
        else:
            length = objref.Length
    return length

def GetEntityNormal(objid:ObjectId):
    with transaction() as trans:
        objref = trans.GetObject(objid, OpenMode.ForRead)
        normal = objref.Normal
    return normal


def GetIdListBoundXYZ(objidlist):
    with transaction() as trans:
        extend = Extents3d()
        for objid in objidlist:
            objref = trans.GetObject(objid, OpenMode.ForRead)
            extend.AddExtents(objref.GeometricExtents) 
        point1 = extend.MinPoint
        point2 = extend.MaxPoint
    return [point1.X, point1.Y, point1.Z], [point2.X, point2.Y, point2.Z]

def GetIdListBoundCenterXYZ(objidlist):
    with transaction() as trans:
        extend = Extents3d()
        for objid in objidlist:
            objref = trans.GetObject(objid, OpenMode.ForRead)
            extend.AddExtents(objref.GeometricExtents) 
        point1 = extend.MinPoint
        point2 = extend.MaxPoint
    return [(point1.X+point2.X)/2, (point1.Y+point2.Y)/2, (point1.Z+point2.Z)/2]

def GetIdListBoundXY0(objidlist):
    with transaction() as trans:
        extend = Extents3d()
        for objid in objidlist:
            objref = trans.GetObject(objid, OpenMode.ForRead)
            extend.AddExtents(objref.GeometricExtents) 
        point1 = extend.MinPoint
        point2 = extend.MaxPoint
    return [point1.X, point1.Y, 0], [point2.X, point2.Y, 0]

def GetIdListBoundCenterXY0(objidlist):
    with transaction() as trans:
        extend = Extents3d()
        for objid in objidlist:
            objref = trans.GetObject(objid, OpenMode.ForRead)
            extend.AddExtents(objref.GeometricExtents) 
        point1 = extend.MinPoint
        point2 = extend.MaxPoint
    return [(point1.X+point2.X)/2, (point1.Y+point2.Y)/2, 0]



def GetRefListBoundXYZ(objreflist):
    with transaction() as trans:
        extend = Extents3d()
        for objref in objreflist:
            extend.AddExtents(objref.GeometricExtents) 
        point1 = extend.MinPoint
        point2 = extend.MaxPoint
    return [point1.X, point1.Y, point1.Z], [point2.X, point2.Y, point2.Z]


def GetRefListBoundCenterXYZ(objreflist):
    with transaction() as trans:
        extend = Extents3d()
        for objref in objreflist:
            extend.AddExtents(objref.GeometricExtents) 
        point1 = extend.MinPoint
        point2 = extend.MaxPoint
    return [(point1.X+point2.X)/2, (point1.Y+point2.Y)/2, (point1.Z+point2.Z)/2]

def GetRefListBoundXY0(objreflist):
    with transaction() as trans:
        extend = Extents3d()
        for objref in objreflist:
            extend.AddExtents(objref.GeometricExtents) 
        point1 = extend.MinPoint
        point2 = extend.MaxPoint
    return [point1.X, point1.Y, 0], [point2.X, point2.Y, 0]


def GetRefListBoundCenterXY0(objreflist):
    with transaction() as trans:
        extend = Extents3d()
        for objref in objreflist:
            extend.AddExtents(objref.GeometricExtents) 
        point1 = extend.MinPoint
        point2 = extend.MaxPoint
    return [(point1.X+point2.X)/2, (point1.Y+point2.Y)/2, 0]




def GetSSBoundXYZ(ss1:SelectionSet):
    with transaction() as trans:
        extend = Extents3d()
        for objid in ss1.GetObjectIds():
            objref = trans.GetObject(objid, OpenMode.ForRead)
            extend.AddExtents(objref.GeometricExtents) 
        point1 = extend.MinPoint
        point2 = extend.MaxPoint
    return [point1.X, point1.Y, point1.Z], [point2.X, point2.Y, point2.Z]


def GetSSBoundCenterXYZ(ss1:SelectionSet):
    with transaction() as trans:
        extend = Extents3d()
        for objid in ss1.GetObjectIds():
            objref = trans.GetObject(objid, OpenMode.ForRead)
            extend.AddExtents(objref.GeometricExtents) 
        point1 = extend.MinPoint
        point2 = extend.MaxPoint
    return [(point1.X+point2.X)/2, (point1.Y+point2.Y)/2, (point1.Z+point2.Z)/2]



def GetSSBoundXY0(ss1:SelectionSet):
    with transaction() as trans:
        extend = Extents3d()
        for objid in ss1.GetObjectIds():
            objref = trans.GetObject(objid, OpenMode.ForRead)
            extend.AddExtents(objref.GeometricExtents) 
        point1 = extend.MinPoint
        point2 = extend.MaxPoint
    return [point1.X, point1.Y, 0], [point2.X, point2.Y, 0]



def GetSSBoundCenterXY0(ss1:SelectionSet):
    with transaction() as trans:
        extend = Extents3d()
        for objid in ss1.GetObjectIds():
            objref = trans.GetObject(objid, OpenMode.ForRead)
            extend.AddExtents(objref.GeometricExtents) 
        point1 = extend.MinPoint
        point2 = extend.MaxPoint
    return [(point1.X+point2.X)/2, (point1.Y+point2.Y)/2, 0]

def GetMKPolyLinePointList(objid:ObjectId):
    result = []
    with transaction() as trans:
        pline = trans.GetObject(objid, OpenMode.ForRead)            
        if pline.Closed:
            for i in range(pline.NumberOfVertices):
                point = pline.GetPoint3dAt(i)
                result.append([point.X, point.Y, point.Z])
    return result

def GetLWPolyLinePointList(objid:ObjectId, pt1=None):
    with transaction() as trans:
        pline = trans.GetObject(objid, OpenMode.ForRead)
        po1 = [pline.StartPoint.X, pline.StartPoint.Y, pline.StartPoint.Z]
        po2 = [pline.EndPoint.X, pline.EndPoint.Y, pline.EndPoint.Z]
        result = []
        for i in range(pline.NumberOfVertices):
            point = pline.GetPoint3dAt(i)
            result.append([point.X, point.Y, point.Z])
    if pline.Closed: result = result + result[0:1]
    if pt1 == None: return result
    if Distance(pt1, po1) > Distance(pt1, po2): result = result[::-1]
    return result


def GetLWPolyLineLengthAndLengthAtPoint(objid:ObjectId, pt1):
    with transaction() as trans:
        objref = trans.GetObject(objid, OpenMode.ForRead)
        length = objref.GetDistAtPoint(ToPoint3d(pt1)) 
    return objref.Length, length



def GetMKPolyLineDirectList(objid:ObjectId):
    ptlist = GetMKPolyLinePointList(objid)
    if ptlist == []: return []
    ptlist = ptlist + ptlist[0:1]
    drlist = []
    for i in range(len(ptlist)-1):
        pt1 = ptlist[i]
        pt2 = ptlist[i+1]
        dr1 = Direct(pt1, pt2)
        drlist.append(dr1)
    return drlist


def GetLWPolyLineDirectList(objid:ObjectId, pt1=None):
    ptlist = GetLWPolyLinePointList(objid, pt1)
    drlist = []
    for i in range(len(ptlist)-1):
        pt1 = ptlist[i]
        pt2 = ptlist[i+1]
        dr1 = Direct(pt1, pt2)
        drlist.append(dr1)
    return drlist

def GetLWPolyLineLengthList(objid:ObjectId, pt1=None):
    ptlist = GetLWPolyLinePointList(objid, pt1)
    lengthlist = []
    for i in range(len(ptlist)-1):
        pt1 = ptlist[i]
        pt2 = ptlist[i+1]
        length = Distance(pt1, pt2)
        lengthlist.append(length)
    return lengthlist


def GetLWPolyLineMidPointList(objid:ObjectId, pt1=None):
    pline_point_list = GetLWPolyLinePointList(objid, pt1)
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

llzhu_entnet_last = None
def SetEntNext():
    global llzhu_entnet_last
    llzhu_entnet_last = Utils.EntLast()

def GetEntNextIdList():
    global llzhu_entnet_last
    objid = llzhu_entnet_last
    objidlist = []
    while True:
        objid = Utils.EntNext(objid, skipSubEnt=True)
        if objid.IsNull: break
        objidlist.append(objid)
    llzhu_entnet_last = None
    return objidlist 


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

def CommandAddPoint(pt1):
    Command(["point", ToPoint3d(pt1)]), Prompt("\n")
    
def CommandAddText(pt1, string, size): # pt1 = [x, y, z]
    Command(["text", ToPoint3d(pt1), System.Int32(size), System.Int32(0), string]), Prompt("\n")
    
def CommandAddLine(pt1, pt2): # pt1 = [x, y, z]
    Command(["LINE", ToPoint3d(pt1), ToPoint3d(pt2), ""]), Prompt("\n")
    
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
    

# def CommandAddSweep()



def CommandZoom(pt1, pt2):
    Command(["zoom", ToPoint3d(pt1), ToPoint3d(pt2), ""]), Prompt("\n")


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
    Command(["rotate", objid, "", ToPoint3d(centert_point), System.Double(angle)]), Prompt("\n")
    


def CommandRotate3d(objid:ObjectId|SelectionSet, pt1, pt2, angle):
    Command(["rotate3d", objid, "", ToPoint3d(pt1), ToPoint3d(pt2), System.Double(angle)]), Prompt("\n")


def CommandRotateCopy(objid:ObjectId|SelectionSet, centert_point, angle):
    Command(["rotate", objid, "", ToPoint3d(centert_point), "C", System.Double(angle)]), Prompt("\n")
    
def CommandErase(objid:ObjectId|SelectionSet):
    Command(["erase", objid]), Prompt("\n")

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

zhu_count = 0
def CountColorIndex(): # [1:254] # 0 layer # 255 bloclk
    global zhu_count
    zhu_count +=1 
    if zhu_count >= 255: zhu_count = 1
    return zhu_count

def CountColorIndexReset():
    global zhu_count
    zhu_count = 0

def CountColorIndexSet(count=1):
    global zhu_count
    count = count - 1
    if count < 0: count = 0
    if count > 255: count = 255
    zhu_count = count


def TransExplodeIdList(objidlist):
    buflist = []
    for objid in objidlist:
        objref = trans.GetObject(objid, OpenMode.ForWrite)   
        match str(objref): 
            case "Autodesk.AutoCAD.DatabaseServices.Polyline": # Polyline Rotation3D 后 还是 Polyline
                result = DBObjectCollection()
                objref.Explode(result) # Line Explode 会出错，要求PL or block # Autodesk.AutoCAD.DatabaseServices.Line
                for objref in result: 
                    buflist.append(objref.ObjectId)
            case "Autodesk.AutoCAD.DatabaseServices.BlockReference":
                pass  
            case "Autodesk.AutoCAD.DatabaseServices.Line":
                if objref.Layer != "打标1": buflist.append(objref.ObjectId)
    return buflist

def TransAutoExplodeObjectIdList(objidlist):
    TransExplodeIdListWithErase(objidlist)
def TransExplodeIdListWithErase(objidlist):
    buflist = []
    for objid in objidlist:
        objref = trans.GetObject(objid, OpenMode.ForWrite)   
        match str(objref): 
            case "Autodesk.AutoCAD.DatabaseServices.Polyline": # Polyline Rotation3D 后 还是 Polyline
                result = DBObjectCollection()
                objref.Explode(result) # Line Explode 会出错，要求PL or block # Autodesk.AutoCAD.DatabaseServices.Line
                objref.Erase()
                for objref in result: 
                    currentblock.AppendEntity(objref) # objid = currentblock.AppendEntity(objref)
                    trans.AddNewlyCreatedDBObject(objref, True)
                    buflist.append(objref.ObjectId)
            case "Autodesk.AutoCAD.DatabaseServices.BlockReference":
                pass  
            case "Autodesk.AutoCAD.DatabaseServices.Line":
                if objref.Layer != "打标1": buflist.append(objref.ObjectId)
    return buflist



def TransAutoFindRegionRectList(objidlist):
    objreflist = []
    for objid in objidlist:
        objref = trans.GetObject(objid, OpenMode.ForRead)
        matrix4x4 = Matrix3d.Identity # 单位矩阵
        objrefcopy = objref.GetTransformedCopy(matrix4x4)
        objreflist.append(objrefcopy)
        
    collect = DBObjectCollection()
    for objref in objreflist:
        collect.Add(objref)

    regions = Region.CreateFromCurves(collect)
    buflist = []
    for region in regions:
        area = region.Area
        extend = region.GeometricExtents
        pt1, pt2 = [extend.MinPoint.X, extend.MinPoint.Y, extend.MinPoint.Z], [extend.MaxPoint.X, extend.MaxPoint.Y, extend.MaxPoint.Z]
        center = [(extend.MinPoint.X+extend.MaxPoint.X)/2, (extend.MinPoint.Y+extend.MaxPoint.Y)/2, (extend.MinPoint.Z+extend.MaxPoint.Z)/2]
        buflist.append([pt1, pt2, center, area])
    buflist.sort(key = lambda item: item[3], reverse=True) # 排序规则，reverse = True 降序， reverse = False 升序（默认）

    # 面积移除
    count = len(buflist)
    subindexlist = []
    for i in range(count-1):
        if i in subindexlist: continue
        pt1, pt2, center1, area1  = buflist[i]
        x1, y1 = pt1[0:2]
        x2, y2 = pt2[0:2]
        pt1 = [x1, y1]
        pt2 = [x2, y1]
        pt3 = [x2, y2]
        pt4 = [x1, y2]
        ptlist = [pt1, pt2, pt3, pt4]
        for m in range(i+1, count):
            po1, po2, center2, area2  = buflist[m]
            if IsPointInRect(center2, ptlist): subindexlist.append(m)

    result = []
    for i, [pt1, pt2, center, area] in enumerate(buflist):
        if i in subindexlist: continue
        result.append([pt1, pt2]) 
    return result

def TransAutoFindRectFencePointList(pb1, pb2, objidlist):
    result = []
    for objid in objidlist:
        objref = trans.GetObject(objid, OpenMode.ForRead)
        extend = objref.GeometricExtents
        center = [(extend.MinPoint.X+extend.MaxPoint.X)/2, (extend.MinPoint.Y+extend.MaxPoint.Y)/2, 0]
        ptlist = TransLWPolyLinePointList(objid)
        for i in range(len(ptlist)-1):
            pt1 = ptlist[i]
            pt2 = ptlist[i+1]
            perflag1 = GetPerflagXY(pb1, pb2, pt1)
            perflag2 = GetPerflagXY(pb1, pb2, pt2)
            if perflag1 >= 0 and perflag2 >= 0: continue
            if perflag1 <= 0 and perflag2 <= 0: continue
            perflag3 = GetPerflagXY(pt1, pt2, pb1)
            perflag4 = GetPerflagXY(pt1, pt2, pb2)
            if perflag3 >= 0 and perflag4 >= 0: continue
            if perflag3 <= 0 and perflag4 <= 0: continue

            perflag0 = GetPerflagXY(pt1, pt2, center)
            perflag0 = -perflag0
            direct = GetPerDirectWithPerflagXY(pt1, pt2, perflag0)
            result.append([pt1, pt2, direct]) 
    return result



def TransFenceLine2PointList(objidlist):
    result = []
    objid1, objid2 = objidlist[0:2]
    pt1, mid1, pt2 = TransEntityStartMidEndPoint(objid1)
    po1, mid2, po2 = TransEntityStartMidEndPoint(objid2)
    dr1 = Direct(mid2, mid1)
    dr2 = Direct(mid1, mid2)
    result.append([pt1, pt2, dr1])
    result.append([po1, po2, dr2])
    return result


def TransAutoFindRectPointList(objidlist):
    result = []
    for objid in objidlist:
        objref = trans.GetObject(objid, OpenMode.ForRead)
        extend = objref.GeometricExtents
        center = [(extend.MinPoint.X+extend.MaxPoint.X)/2, (extend.MinPoint.Y+extend.MaxPoint.Y)/2, 0]
        ptlist = TransLWPolyLinePointList(objid)
        for i in range(len(ptlist)-1):
            pt1 = ptlist[i]
            pt2 = ptlist[i+1]
            pt3 = center
            perflag = GetPerflagXY(pt1, pt2, pt3)
            perflag = -perflag
            direct = GetPerDirectWithPerflagXY(pt1, pt2, perflag) 
            result.append([pt1, pt2, direct]) 
    return result



def TransAutoDimTextAlignPointList(objidlist, pt1, pt2):
    dr1 = Direct(pt1, pt2)
    angle1 = AngleFromDotDr1Dr2(dr1, [ 1, 0, 0])
    angle2 = AngleFromDotDr1Dr2(dr1, [-1, 0, 0])
    angle3 = AngleFromDotDr1Dr2(dr1, [0,  1, 0])
    angle4 = AngleFromDotDr1Dr2(dr1, [0, -1, 0])
    if angle1 < 45: flagd = "+X"
    if angle2 < 45: flagd = "-X"
    if angle3 < 45: flagd = "+Y"
    if angle4 < 45: flagd = "-Y"

    auflist = []
    for objid in objidlist:
        objref = trans.GetObject(objid, OpenMode.ForRead)
        # po0 = [objref.TextPosition.X, objref.TextPosition.Y, objref.TextPosition.Z]
        po1, po2 = TransDimTextBoundXY0(objid)
        match flagd:
            case "+X"|"-X": po0 = po1[:]
            case "+Y"|"-Y": po0 = po2[:]
        # pd1 = [objref.XLine1Point.X, objref.XLine1Point.Y, objref.XLine1Point.Z] 
        # pd2 = [objref.XLine2Point.X, objref.XLine2Point.Y, objref.XLine2Point.Z] 
        # pd3 = [objref.DimLinePoint.X, objref.DimLinePoint.Y, objref.DimLinePoint.Z] 
        dimmid = [(objref.XLine1Point.X + objref.XLine2Point.X)/2, (objref.XLine1Point.Y + objref.XLine2Point.Y)/2, 0]
        auflist.append([objid, po0, po1, po2, dimmid])

    

    match flagd:
        case "+X":
            buflist = []
            for objid, po0, po1, po2, dimmid in auflist:
                x1, y1, z1 = dimmid
                buflist.append([x1, objid, po0, po1, po2, dimmid])
            buflist.sort(key = lambda item: item[0], reverse=False)  # 小 -> 大

            for i in range(1, len(buflist)):
                aa, objid, po0, po1, po2, dimmid1 = buflist[i-1]
                bb, objid, po0, pt1, pt2, dimmid2 = buflist[i]
                x1, y1, z1 = po1
                x2, y2, z2 = po2
                length1, pad = x2 - x1, y2 - y1
                x3, y3, z3 = pt1
                x4, y4, z4 = pt2
                length2, pad = x4 - x3, y4 - y3
                xd, yd, zd = dimmid2

                xt1 = x1 + length1 + pad
                xt2 = xd + 0.1
                xt = max(xt1, xt2)
                dx = xt - x3
                buflist[i][3] = Vec3Add(pt1, [dx, 0, 0])
                buflist[i][4] = Vec3Add(pt2, [dx, 0, 0])

            y0 = buflist[0][2][1]
            for i in range(len(buflist)):
                buflist[i][3][1] = y0
                buflist[i][4][1] = y0
        case "-X":
            buflist = []
            for objid, po0, po1, po2, dimmid in auflist:
                x1, y1, z1 = dimmid
                buflist.append([x1, objid, po0, po1, po2, dimmid])
            buflist.sort(key = lambda item: item[0], reverse=True)  # 大 -> 小

            for i in range(1, len(buflist)):
                aa, objid, po0, po1, po2, dimmid1 = buflist[i-1]
                bb, objid, po0, pt1, pt2, dimmid2 = buflist[i]
                x1, y1, z1 = po1
                x2, y2, z2 = po2
                length1, pad = x2 - x1, y2 - y1
                x3, y3, z3 = pt1
                x4, y4, z4 = pt2
                length2, pad = x4 - x3, y4 - y3
                xd, yd, zd = dimmid2

                xt1 = x1 - pad - length2
                xt2 = xd - 0.1 - length2
                xt = min(xt1, xt2)
                dx = xt - x3
                buflist[i][3] = Vec3Add(pt1, [dx, 0, 0])
                buflist[i][4] = Vec3Add(pt2, [dx, 0, 0])

            y0 = buflist[0][2][1]
            for i in range(len(buflist)):
                buflist[i][3][1] = y0
                buflist[i][4][1] = y0


        case "+Y":
            buflist = []
            for objid, po0, po1, po2, dimmid in auflist:
                x1, y1, z1 = dimmid
                buflist.append([y1, objid, po0, po1, po2, dimmid])
            buflist.sort(key = lambda item: item[0], reverse=False)  # 小 -> 大

            for i in range(1, len(buflist)):
                aa, objid, po0, po1, po2, dimmid1 = buflist[i-1]
                bb, objid, po0, pt1, pt2, dimmid2 = buflist[i]
                x1, y1, z1 = po1
                x2, y2, z2 = po2
                pad, length1 = x2 - x1, y2 - y1
                x3, y3, z3 = pt1
                x4, y4, z4 = pt2
                pad, length2 = x4 - x3, y4 - y3
                xd, yd, zd = dimmid2

                yt1 = y1 + length1 + pad
                yt2 = yd + 0.1
                yt = max(yt1, yt2)
                dy = yt - y3
                buflist[i][3] = Vec3Add(pt1, [0, dy, 0])
                buflist[i][4] = Vec3Add(pt2, [0, dy, 0])

            x0 = buflist[0][2][0]
            for i in range(len(buflist)):
                buflist[i][3][0] = x0
                buflist[i][4][0] = x0

        case "-Y":
            buflist = []
            for objid, po0, po1, po2, dimmid in auflist:
                x1, y1, z1 = dimmid
                buflist.append([y1, objid, po0, po1, po2, dimmid])
            buflist.sort(key = lambda item: item[0], reverse=True)  # 大 -> 小

            for i in range(1, len(buflist)):
                aa, objid, po0, po1, po2, dimmid1 = buflist[i-1]
                bb, objid, po0, pt1, pt2, dimmid2 = buflist[i]
                x1, y1, z1 = po1
                x2, y2, z2 = po2
                pad, length1 = x2 - x1, y2 - y1
                x3, y3, z3 = pt1
                x4, y4, z4 = pt2
                pad, length2 = x4 - x3, y4 - y3
                xd, yd, zd = dimmid2

                yt1 = y1 - pad - length2
                yt2 = yd - 0.1 - length2
                yt = min(yt1, yt2)
                dy = yt - y3
                buflist[i][3] = Vec3Add(pt1, [0, dy, 0])
                buflist[i][4] = Vec3Add(pt2, [0, dy, 0])

            x0 = buflist[0][2][0]
            for i in range(len(buflist)):
                buflist[i][3][0] = x0
                buflist[i][4][0] = x0

    result = []
    for aa, objid, po0, po1, po2, dimmid1 in buflist:

        if Distance(dimmid1, po1) < Distance(dimmid1, po2): 
            result.append([objid, po1])
        else:
            result.append([objid, po2])

    return result



def TransAutoDimTextAlignLayer2PointList(objidlist, pt1, pt2):
    dr1 = Direct(pt1, pt2)
    angle1 = AngleFromDotDr1Dr2(dr1, [ 1, 0, 0])
    angle2 = AngleFromDotDr1Dr2(dr1, [-1, 0, 0])
    angle3 = AngleFromDotDr1Dr2(dr1, [0,  1, 0])
    angle4 = AngleFromDotDr1Dr2(dr1, [0, -1, 0])
    if angle1 < 45: flagd = "+X"
    if angle2 < 45: flagd = "-X"
    if angle3 < 45: flagd = "+Y"
    if angle4 < 45: flagd = "-Y"

    auflist = []
    for i, objid in enumerate(objidlist):
        objref = trans.GetObject(objid, OpenMode.ForRead)
        # po0 = [objref.TextPosition.X, objref.TextPosition.Y, objref.TextPosition.Z]
        po1, po2 = TransDimTextBoundXY0(objid)
        x1, y1, z1 = po1
        x2, y2, z2 = po2
        numb = i%2 +0.5
        match flagd:
            case "+X": 
                length, pad = x2 - x1, y2 - y1
                dr = [0, numb*pad, 0]
                po0 = Vec3Add(po1, dr)
            case "-X": 
                length, pad = x2 - x1, y2 - y1
                dr = [0, -numb*pad, 0]
                po0 = Vec3Add(po1, dr)
            case "+Y": 
                pad, length = x2 - x1, y2 - y1
                dr = [numb*pad, 0, 0]
                po0 = Vec3Add(po2, dr)
            case "-Y": 
                pad, length = x2 - x1, y2 - y1
                dr = [-numb*pad, 0, 0]
                po0 = Vec3Add(po2, dr)
        # pd1 = [objref.XLine1Point.X, objref.XLine1Point.Y, objref.XLine1Point.Z] 
        # pd2 = [objref.XLine2Point.X, objref.XLine2Point.Y, objref.XLine2Point.Z] 
        # pd3 = [objref.DimLinePoint.X, objref.DimLinePoint.Y, objref.DimLinePoint.Z] 
        dimmid = [(objref.XLine1Point.X + objref.XLine2Point.X)/2, (objref.XLine1Point.Y + objref.XLine2Point.Y)/2, 0]
        auflist.append([objid, po0, po1, po2, dimmid])


    result = []
    for aa, objid, po0, po1, po2, dimmid1 in auflist:
        result.append([objid, po0])

    return result




def TransAutoPt1pt2ListToMKPolyLine(pt1pt2list, layer_name="排版1", color_index=0):
    count  = len(pt1pt2list)
    buflist = pt1pt2list[0:1]
    reflist = pt1pt2list[1:]
    indexlist = []
    for i in range(count-1):
        pt1, pt2 = buflist[-1]
        for k, [po1, po2] in enumerate(reflist):
            if k in indexlist: continue
            flag = False
            if IsPointSame(pt2, po1):
                flag = True
                buflist.append([po1, po2])
            if IsPointSame(pt2, po2):
                flag = True
                buflist.append([po2, po1])
            if flag:
                indexlist.append(k)
                break
    ptlist = [pt1 for [pt1, pt2] in  buflist]
    return AddMKPolyLine(ptlist, layer_name, color_index)

def TransAutoFindMidLWPolyLine(objid1, objid2):
    objref1 = trans.GetObject(objid1, OpenMode.ForRead)
    pt1 = [objref1.StartPoint.X, objref1.StartPoint.Y, objref1.StartPoint.Z]
    pt2 = [objref1.EndPoint.X, objref1.EndPoint.Y, objref1.EndPoint.Z]
    objref2 = trans.GetObject(objid2, OpenMode.ForRead)
    po1 = [objref2.StartPoint.X, objref2.StartPoint.Y, objref2.StartPoint.Z]
    po2 = [objref2.EndPoint.X, objref2.EndPoint.Y, objref2.EndPoint.Z]
    length1, length2 = Distance(po1, pt1), Distance(po1, pt2)
    pd1 = pt1
    if length1 > length2: pd1 = pt2
    ness = min(length1, length2)
    ness_half = ness / 2 
    objref3 = objref2.GetOffsetCurves( ness_half)[0] # collection
    objref4 = objref2.GetOffsetCurves(-ness_half)[0] # collection
    pk1 = [objref3.StartPoint.X, objref3.StartPoint.Y, objref3.StartPoint.Z]
    pm1 = [objref4.StartPoint.X, objref4.StartPoint.Y, objref4.StartPoint.Z]
    if Distance(pd1, pk1) < Distance(pd1, pm1): return AddDBObject(objref3), ness
    return AddDBObject(objref4), ness

# def GeometryExternalCurve3dToDBLine(curve): 
#     BrepCurveToDBLine(curve)
# def BrepExternalCurve3dToDBLine(curve): 
#     BrepCurveToDBLine(curve)
def BrepCurveToDBLine(curve): # ExternalCurve3d
    return Line(curve.NativeCurve.StartPoint, curve.NativeCurve.EndPoint) # 

 
# def GeometryExternalCurve3dToDBArc(curve): 
#     BrepCurveToDBArc(curve)
# def BrepExternalCurve3dToDBArc(curve): 
#     BrepCurveToDBArc(curve)
def BrepCurveToDBArc(curve): # ExternalCurve3d
    return Arc(curve.NativeCurve.Center, curve.NativeCurve.Normal,curve.NativeCurve.Radius,curve.NativeCurve.StartAngle, curve.NativeCurve.EndAngle)
# Arc(Point3d center, Vector3d normal,double radius,double startAngle, double endAngle)
# Arc(Point3d center, double radius, double startAngle, double endAngle)




def DBObjectConvertRegionToPolylineXY0(objref):
    resultlist = []
    brep = Brep(objref)
    for face in brep.Faces: 
        for loop in face.Loops: # 已经自带排序了
            buflist = [None]
            for i, edge in enumerate(loop.Edges): 
                buflist.append(None)
                if edge.Curve.IsLineSegment:
                    point2d1 = Point2d(edge.Curve.NativeCurve.StartPoint.X, edge.Curve.NativeCurve.StartPoint.Y)
                    point2d2 = Point2d(edge.Curve.NativeCurve.EndPoint.X, edge.Curve.NativeCurve.EndPoint.Y)
                    buflist[i] = [point2d1, 0, 0, 0]
                    buflist[i+1] = [point2d2, 0, 0, 0]
                if edge.Curve.IsCircularArc: 
                    point2d1 = Point2d(edge.Curve.NativeCurve.StartPoint.X, edge.Curve.NativeCurve.StartPoint.Y)
                    point2d2 = Point2d(edge.Curve.NativeCurve.EndPoint.X, edge.Curve.NativeCurve.EndPoint.Y)
                    if IsPointSame([point2d1.X, point2d1.Y, 0], [point2d2.X, point2d2.Y, 0]): continue # 跳过圆对象 # 圆弧和圆在region or brep中是一样的对象，区别在圆的起始结束是同一个点
                    R = edge.Curve.NativeCurve.Radius
                    L = point2d1.GetDistanceTo(point2d2)
                    H = R - math.sqrt(R*R - L*L/4)
                    buflist[i] = [point2d1, 2*H/L, 0, 0]
                    buflist[i+1] = [point2d2, 0, 0, 0]
            if buflist == []: continue
            polyline = Polyline()
            for i, [point, bulge, startwidth, endwidth] in buflist:     
                polyline.AddVertexAt(i, point, bulge, startwidth, endwidth)
            resultlist.append(polyline)
    buflist = []
    for objref in resultlist: buflist.append([objref.Area, objref])
    buflist.sort(key = lambda item: item[0], reverse=True) # 从大到小

    resultlist = []
    for area, objref in buflist: resultlist.append(objref)

    return resultlist




def DBObjectConvertRegionToPolylineXYZ(objref):
    resultlist = []
    brep = Brep(objref)
    for face in brep.Faces: 
        for loop in face.Loops: # 已经自带排序了
            buflist = []
            for i, edge in enumerate(loop.Edges): 
                if edge.Curve.IsLineSegment:
                    polyline = Polyline()
                    point2d1 = Point2d(edge.Curve.NativeCurve.StartPoint.X, edge.Curve.NativeCurve.StartPoint.Y)
                    point2d2 = Point2d(edge.Curve.NativeCurve.EndPoint.X, edge.Curve.NativeCurve.EndPoint.Y)
                    pt1, pt2 = [point2d1.X, point2d1.Y, 0], [point2d2.X, point2d2.Y, 0]
                    polyline.AddVertexAt(0, point2d1, 0, 0, 0)
                    polyline.AddVertexAt(1, point2d2, 0, 0, 0)
                    buflist.append(polyline)
                if edge.Curve.IsCircularArc: 
                    polyline = Polyline()
                    point2d1 = Point2d(edge.Curve.NativeCurve.StartPoint.X, edge.Curve.NativeCurve.StartPoint.Y)
                    point2d2 = Point2d(edge.Curve.NativeCurve.EndPoint.X, edge.Curve.NativeCurve.EndPoint.Y)
                    pt1, pt2 = [point2d1.X, point2d1.Y, 0], [point2d2.X, point2d2.Y, 0]
                    if IsPointSame(pt1, pt2): continue # 跳过圆对象 # 圆弧和圆在region or brep中是一样的对象，区别在圆的起始结束是同一个点
                    R = edge.Curve.NativeCurve.Radius
                    L = point2d1.GetDistanceTo(point2d2)
                    H = R - math.sqrt(R*R - L*L/4)
                    polyline.AddVertexAt(0, point2d1, 2*H/L, 0, 0)
                    polyline.AddVertexAt(1, point2d2, 0, 0, 0)
                    buflist.append(polyline)
            if buflist == []: continue

            for poly in buflist[1:]: 
                CheckLayerAndColor(poly, "补偿1", 35)
                currentblock.AppendEntity(poly)
                trans.AddNewlyCreatedDBObject(poly, True) # 不用先提交到CAD，可以Explode 
                buflist[0].JoinEntity(poly) # 相连也会出错，原因未知 # JoinEntity下一个对象如果不与现在对象相连，会出错 
            resultlist.append(buflist[0])
    buflist = []
    for objref in resultlist: buflist.append([objref.Area, objref])
    buflist.sort(key = lambda item: item[0], reverse=True) # 从大到小

    resultlist = []
    for area, objref in buflist: resultlist.append(objref)

    return resultlist



      



def DBObjectConvertLineToPolylineXY0(objref):
    polyline = Polyline()
    polyline.AddVertexAt(0, Point2d(objref.StartPoint.X, objref.StartPoint.Y), 0, 0, 0)
    polyline.AddVertexAt(1, Point2d(objref.EndPoint.X, objref.EndPoint.Y), 0, 0, 0)
    return polyline

def DBObjectConvertArcToPolylineXY0(objref):
    point1, point2 = Point2d(objref.StartPoint.X, objref.StartPoint.Y), Point2d(objref.EndPoint.X, objref.EndPoint.Y)
    R = objref.Radius
    L = point1.GetDistanceTo(point2)
    H = R - math.sqrt(R*R - L*L/4)
    polyline = Polyline()
    polyline.AddVertexAt(0, Point2d(objref.StartPoint.X, objref.StartPoint.Y), 2*H/L, 0, 0)
    polyline.AddVertexAt(1, Point2d(objref.EndPoint.X, objref.EndPoint.Y), 0, 0, 0)    
    return polyline


def DBObjectConvertCircleToPolylineXY0(objref):
    C = objref.Center
    R = objref.Radius
    point1, point2 = Point2d(C.X+R, C.Y), Point2d(C.X-R, C.Y)
    polyline = Polyline()
    polyline.AddVertexAt(0, point1, 1, 0, 0)
    polyline.AddVertexAt(1, point2, 1, 0, 0)
    polyline.AddVertexAt(2, point1, 1, 0, 0)
    # polyline.Color = reg.Color
    # polyline.LineWeight = reg.LineWeight
    # polyline.LinetypeId = reg.LinetypeId
    return polyline



def DBObjectConvertPolylineToPolylineXY0(objref):
    polyline = Polyline()
    for i in range(objref.NumberOfVertices):
        polyline.AddVertexAt(i, Point2d(objref.GetPoint3dAt(i).X, objref.GetPoint3dAt(i).Y), objref.GetBulgeAt(i), objref.GetStartWidthAt(i), objref.GetEndWidthAt(i))
    return polyline 






# result = DBObjectCollection()
# objref.Explode(result) # Line Explode 会出错，要求PL or block # Autodesk.AutoCAD.DatabaseServices.Line
# objref.Erase()
# for objref in result: 
#     currentblock.AppendEntity(objref) # objid = currentblock.AppendEntity(objref)
#     trans.AddNewlyCreatedDBObject(objref, True)
#     buflist.append(objref.ObjectId)


# # 使无序Loop变有序
# print(buflist)
# count = len(buflist)
# sortlist = buflist[0:1]
# indexlist = [0]
# for i in range(count):
#     for k, [polyline2, po1, po2] in enumerate(buflist):  
#         if k in indexlist: continue
#         polyline1, pt1, pt2 = sortlist[-1]
#         if IsPointSame(pt1, po1) or IsPointSame(pt1, po2) or IsPointSame(pt2, po1) or IsPointSame(pt2, po2):
#             indexlist.append(k)
#             sortlist.append([polyline2, po1, po2])
#             break
# print(sortlist)


# JoinEntities不会自动进行相连排序，同样会遇到下个对象不相连报错
# entlist = System.Array[Entity](len(buflist)-1) # 'Entity[]' 
# i = 0
# for poly in buflist[1:]:
#     entlist[i] = poly
#     i += 1
# buflist[0].JoinEntities(entlist)
# for poly in buflist[1:]: 
#     CheckLayerAndColor(poly, "补偿1", 35)
#     currentblock.AppendEntity(poly)
#     trans.AddNewlyCreatedDBObject(poly, True) # 不用先提交到CAD，可以Explode 
#     buflist[0].JoinEntity(poly) # JoinEntity下一个对象如果不与现在对象相连，会出错


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


    # buflist = []
    # pb1, pb2 = TransDimObjectIdListBoundXY0(objidlist)
    # flagx = (pb1[0]+pb2[0])/2
    # for objid in objidlist:
    #     objref =  trans.GetObject(objid, OpenMode.ForRead)
    #     pt1, pt2 = TransDimTextBoundXY0(objid)
    #     buflist.append([pt1[0], objid, pt1, pt2])

    # auflist, cuflist = [], []
    # x1, objid, pt1, pt2 = buflist[0]
    # basex, basey, basez = pt1
    # for x1, objid, pt1, pt2 in buflist:
    #     if abs(x1-basex) < 0.0001 and x1 < flagx: auflist.append([x1, objid, pt1, pt2])
    #     if abs(x1-basex) < 0.0001 and x1 > flagx: cuflist.append([x1, objid, pt1, pt2])

    # auflist.sort(key = lambda item: item[0], reverse=True)  # 大 -> 小
    # cuflist.sort(key = lambda item: item[0], reverse=False) # 小 -> 大
    # sumx = auflist[0][0]
    # for i, [x1, objid, pt1, pt2] in enumerate(auflist[1:]):
    #     length, width = pt2[0]-pt1[0], pt2[1]-pt1[1]
    #     sumx -= (length + width)
    #     auflist[i][0] = sumx
    # x1, objid, pt1, pt2 = auflist[0]
    # sumx = x1 + pt2[0]-pt1[0] + pt2[1]-pt1[1]
    # for i, [x1, objid, pt1, pt2] in enumerate(cuflist):
    #         auflist[i][0] = sumx
    #         length, width = pt2[0]-pt1[0], pt2[1]-pt1[1]
    #         sumx += (length + width)

    # result = []
    # for x1, objid, pt1, pt2 in auflist:
    #     result.append(objid, [x1, basey, 0])




# if save: return
# filepath = "F:\\CADdll\\Drawing1.dwg"
# doc.Database.SaveAs(filepath, True, DwgVersion.Current, doc.Database.SecurityParameters)
# doc = Application.DocumentManager.Open(filepath, False)
# Application.DocumentManager.MdiActiveDocument = doc


# def CommandAsyncAddLine(pt1, pt2): # pt1 = [x, y, z] # Error 
#     entcalss.objid = None
#     casync, zhu = CommandAsync(["LINE", ToPoint3d(pt1), ToPoint3d(pt2), ""]), Prompt("\n")
#     casync.OnCompleted(System.Action(__GetEntLastFunc))
#     print("entcalss.objid = ", entcalss.objid)
#     return entcalss.objid


# def GetType(objid):
#     try:
#         typename = str(objid.GetType())
#     except:
#         classname = str(type(objid))
#         typename = classname[8:-2]
#     return typename


#计算 求两个向量的夹角


# 方法1
# 通过两个向量的法向量的点乘的反余弦获取弧度，然后通过弧度获取角度
# Mathf.Acos(Vector3.Dot(a.normal,b.normal))* Mathf.Rad2Deg

# 方法2
# 通过两个向量的法向量的叉乘的模长的反正弦获取弧度，然后通过弧度获取角度
# Mathf.Asin(Vector3.Distance(Vector3.zero,Vector3.Cross(a.normal,b.normal)))* Mathf.Rad2Deg


# def CloseOSNap(mode="节点"):
#     osmode = Application.GetSystemVariable("OSMODE")
#     match mode:
#         case "": Application.SetSystemVariable("OSMODE", System.Int32(ll_old_osmode))

# 0  NON（无）
# 1  END（端点）
# 2  MID（中点）
# 4  CEN（圆心）
# 8  NOD（节点）
# 16  QUA（象限点）
# 32  INT（交点）
# 64  INS（插入点）
# 128  PER（垂足）
# 256  TAN（切点）
# 512  NEA（最近点）
# 1024  QUI（快速）
# 2048  APP（外观交点）
# 4096  EXT（尺寸线）
# 8192  PAR（平行）




# 0  NON（无）
# 1  END（端点）
# 2  MID（中点）
# 4  CEN（圆心）
# 8  NOD（节点）
# 16  QUA（象限点）
# 32  INT（交点）
# 64  INS（插入点）
# 128  PER（垂足）
# 256  TAN（切点）
# 512  NEA（最近点）
# 1024  QUI（快速）
# 2048  APP（外观交点）
# 4096  EXT（尺寸线）
# 8192  PAR（平行）
