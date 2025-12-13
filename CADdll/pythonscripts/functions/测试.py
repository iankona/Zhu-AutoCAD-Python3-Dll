import clr

from Autodesk.AutoCAD.ApplicationServices import Application
from Autodesk.AutoCAD.Runtime import CommandClassAttribute, CommandMethodAttribute

import academit


def 命令(): academit.添加命令("ll-hello3", 函数)
def 函数():
    doc = Application.DocumentManager.MdiActiveDocument
    doc.Editor.WriteMessage("\nHello, python, 鹅鹅鹅")


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
