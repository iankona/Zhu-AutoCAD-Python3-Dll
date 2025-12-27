import clr

import acad
import academit

import System
from Autodesk.AutoCAD.ApplicationServices import Application
# from Autodesk.AutoCAD.EditorInput
# from Autodesk.AutoCAD.Runtime
from Autodesk.AutoCAD.DatabaseServices import Line, ObjectId, Transaction, OpenMode, BlockTable, BlockTableRecord
from Autodesk.AutoCAD.Geometry import Point3d


# doc = Application.DocumentManager.MdiActiveDocument
# doc.Editor.WriteMessage("\nHello, python, 鹅鹅鹅")


# doc = Application.DocumentManager.MdiActiveDocument
# ed = doc.Editor
# db = doc.Database


def 命令(): 
    academit.添加命令("ll-line", ll_line)
    academit.添加命令("ll-pline-point", ll_pline_point)
    academit.添加命令("ll-offset", ll_offset)
    academit.添加命令("ll-get-osmode", 函数2)
    academit.添加命令("ll-set-osmode", 函数3)


@acad.decorator_command
def ll_line():
    with acad.transaction() as trans:
        acad.AddLine([50.5,50], [500,500])
        acad.AddPolyline([[100,100], [100,300], [300,500], [400,600], [100,100]], "图层2", 2)
        # acad.AddRect([1000,1000], [1500,1800])
        # acad.AddCircle([1000,1000], 500)
        # acad.AddCircle([1500,1000], 300)





def ll_pline_point():
    doc, ed, db, Command, CommandAsync = acad.GetActiveDocument()
    objid = acad.EntSel("请选择1条多段线:")
    dlock = doc.LockDocument()
    trans = db.TransactionManager.StartTransaction() # if works !!!
    pline = trans.GetObject(objid, OpenMode.ForRead)
    print(objid, pline.GetType(), pline.StartPoint, pline.EndPoint, pline.NumberOfVertices) # .ToString()
    列表 = []
    for i in range(pline.NumberOfVertices):
        point = pline.GetPoint2dAt(i)
        列表.append(point)
    print(列表)
    dlock.Dispose()




def ll_offset():
    acad.GetActiveDocument()
    objid = acad.EntSel()
    # print("objid:", objid) # (1960788914368)
    acad.OffSet(objid, 53.5, [0,0])
    objid = acad.EntSel()
    acad.OffSet(objid, 76, [0,0], "当前")



def 函数2():
    a = Application.GetSystemVariable("OSMODE")
    print(a)
    # ed.WriteMessage("\n旧捕捉:"+str(a))
    Application.SetSystemVariable("OSMODE", System.Int64(0))

def 函数3():
    # 给系统变量设置新值
    Application.SetSystemVariable("OSMODE", System.Int64(1))


# 端点(Endpoint)捕捉到线段、圆弧、多段线等对象的端点。
# 中点(Midpoint):捕捉到线段、圆弧、多段线等对象的中点。
# 交点(Intersection):捕捉到两个对象的交点。
# 垂足(Perpendicular):换句话说，捕捉到与选定对象垂直的点。
# 切点(Tangent)捕捉到与圆或圆弧相切的点。
# 插人点(Insert)捕捉到块或属性的插入点。
# 最近点(Nearest):捕捉到对象上离光标最近的点。

# 初始值：4133
# 使用以下位码设置“对象捕捉”的运行模式(OSNAP mode)：
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