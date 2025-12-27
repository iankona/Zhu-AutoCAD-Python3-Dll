import clr

from Autodesk.AutoCAD.ApplicationServices import Application
from Autodesk.AutoCAD.Runtime import CommandClassAttribute, CommandMethodAttribute

import acad
import academit


def 命令(): 
    academit.添加命令("ll-hello3", 函数)

@acad.test_decorator2
@acad.test_decorator1
def 函数():
    doc = Application.DocumentManager.MdiActiveDocument
    doc.Editor.WriteMessage("Hello, python, 鹅鹅鹅\n")
    with acad.test_context2(), acad.test_context1():
        doc.Editor.WriteMessage("Hello, autocad, 鹅鹅鹅\n")



# 装饰器2.1之前...
# 装饰器1.1之前...
# Hello, python, 鹅鹅鹅
# 装饰器1.2之后...
# 装饰器2.2之后...


# 装饰器2.1之前...
# 装饰器1.1之前...
# Hello, python, 鹅鹅鹅

# 管理器2.1之前...
# 管理器1.1之前...
# Hello, autocad, 鹅鹅鹅
# 管理器1.2之后...
# 管理器2.2之后...

# 装饰器1.2之后...
# 装饰器2.2之后...



# using System;
# using System.Collections.Generic;
# using System.Text;

# //*******************************************//
# //               Type Library                //
# //*******************************************//
# using Autodesk.AutoCAD.Interop;        // AutoCAD 2008 TypeLibrary
# using Autodesk.AutoCAD.Interop.Common; // AutoCAD/ObjectDBXCommon 17.0 Type Library
# using Autodesk.AutoCAD.Customization;  // accui.dll

# //*******************************************//
# //               acdbmgd.dll                 //
# //*******************************************//
# using Autodesk.AutoCAD.Runtime;
# using Autodesk.AutoCAD.LayerManager;
# using Autodesk.AutoCAD.GraphicsSystem;
# using Autodesk.AutoCAD.GraphicsInterface;
# using Autodesk.AutoCAD.Geometry;
# using Autodesk.AutoCAD.DatabaseServices.Filters;
# using Autodesk.AutoCAD.DatabaseServices;
# using Autodesk.AutoCAD.Colors;

# //******************************************//
# //                 acmgd.dll                //
# //******************************************//
# using Autodesk.AutoCAD.Windows.ToolPalette;
# using Autodesk.AutoCAD.Windows;
# using Autodesk.AutoCAD.Publishing;
# using Autodesk.AutoCAD.PlottingServices;
# using Autodesk.AutoCAD.EditorInput;
# using Autodesk.AutoCAD.ApplicationServices;
