import clr
import math
import acad
import academit

import System

def 命令(): 
    academit.添加命令("llcube", llcube)
    # academit.添加命令("ll-entsel", ll_entsel)

llzhu_cube_length = 100
llzhu_cube_width = 200
llzhu_cube_height = 300
llzhu_cube_show_angle = 30
def llzhu_ui_cube_input():
    global llzhu_cube_length, llzhu_cube_width, llzhu_cube_height, llzhu_cube_show_angle
    length = acad.GetDouble(llzhu_cube_length, "请输入长度: ")
    width = acad.GetDouble(llzhu_cube_width, "请输入宽度: ")
    height = acad.GetDouble(llzhu_cube_height, "请输入高度: ")
    angle = acad.GetDouble(llzhu_cube_show_angle, "请输入显示轴线角度: ")
    if length != None: llzhu_cube_length = length
    if width != None: llzhu_cube_width = width
    if height != None: llzhu_cube_height = height
    if angle != None: llzhu_cube_show_angle = angle

llzhu_dim_height = 100
llzhu_dim_biaohao = 30
def llzhu_ui_dim_input():
    global llzhu_dim_height, llzhu_dim_biaohao
    height = acad.GetInt(llzhu_dim_height, "请输入标注高度:")
    biaohao = acad.GetInt(llzhu_dim_biaohao, "请输入标注标号:")
    if height != None: llzhu_dim_height = height
    if biaohao != None: llzhu_dim_biaohao = biaohao

@acad.decorator_command
def llcube():
    llzhu_ui_cube_input()
    pt0 = acad.GetPoint()
    llzhu_ui_dim_input()
    with acad.transaction() as trans:
        # 绘制立方体
        pt1 = pt0
        pt2 = acad.Vec3Add(pt1, [0, llzhu_cube_height])
        po1 = acad.Vec3Add(pt1, [llzhu_cube_length, 0])
        po2 = acad.Vec3Add(pt1, [llzhu_cube_length, llzhu_cube_height])
        acad.AddLWPolyLine([pt1, pt2, po2, po1, pt1])

        dx = llzhu_cube_width*math.sin(llzhu_cube_show_angle*0.01745329252)
        dy = llzhu_cube_width*math.cos(llzhu_cube_show_angle*0.01745329252)

        pw1 = acad.Vec3Add(pt2, [dx, dy])
        pw2 = acad.Vec3Add(po2, [dx, dy])
        pw3 = acad.Vec3Add(po1, [dx, dy])
        acad.AddLine(pt2, pw1)
        acad.AddLine(po2, pw2)
        acad.AddLine(po1, pw3)
        acad.AddLWPolyLine([pw1, pw2, pw3])
        # 标注立方体

        acad.AddDimLinear(pt1, po1, direct="-y", dle=llzhu_dim_height, dimflagnum=llzhu_dim_biaohao)
        acad.AddDalLinear(pt2, pw1, direct="+y", dle=llzhu_dim_height, dimflagnum=llzhu_dim_biaohao)
        acad.AddDimLinear(pw2, pw3, direct="+x", dle=llzhu_dim_height, dimflagnum=llzhu_dim_biaohao)


    # pt1 = acad.Vec3Add(pt0, [0, -2000-width])
    # # 绘制俯视图
    # pt2 = acad.Vec3Add(pt1, [0, width])
    # po1 = acad.Vec3Add(pt1, [length, 0])
    # po2 = acad.Vec3Add(pt1, [length, width])
    # acad.AddLwpline([pt1, pt2, po2, po1, pt1])
    # acad.AddDimLinear(pt2, po2, direct="+y", dle=dlen)
    # acad.AddDimLinear(pt1, pt2, direct="-x", dle=dlen)


    # pt1 = acad.Vec3Add(pt0, [0, -2000-width-1000-height])
    # # 绘制正视图
    # pt2 = acad.Vec3Add(pt1, [0, height])
    # po1 = acad.Vec3Add(pt1, [length, 0])
    # po2 = acad.Vec3Add(pt1, [length, height])
    # acad.AddLwpline([pt1, pt2, po2, po1, pt1])
    # acad.AddDimLinear(pt2, po2, direct="+y", dle=dlen)
    # acad.AddDimLinear(pt1, pt2, direct="-x", dle=dlen)

    # pt1 = acad.Vec3Add(pt0, [+length+1000, -2000-height])
    # # 绘制侧视图
    # pt2 = acad.Vec3Add(pt1, [0, height])
    # po1 = acad.Vec3Add(pt1, [width, 0])
    # po2 = acad.Vec3Add(pt1, [width, height])
    # acad.AddLwpline([pt1, pt2, po2, po1, pt1])
    # acad.AddDimLinear(pt2, po2, direct="+y", dle=dlen)
    # acad.AddDimLinear(pt1, pt2, direct="-x", dle=dlen)



# def ll_entsel():
#     acad.GetActiveDocument()
#     acad.EntSel()


# using Autodesk.AutoCAD.Runtime;
# using Autodesk.AutoCAD.ApplicationServices;
# using Autodesk.AutoCAD.DatabaseServices;
# using Autodesk.AutoCAD.Geometry;
 
# [CommandMethod("CopyDimStyles")]
# public static void CopyDimStyles()
# {
#     // Get the current database
#     Document acDoc = Application.DocumentManager.MdiActiveDocument;
#     Database acCurDb = acDoc.Database;

#     // Start a transaction
#     using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
#     {
#         // Open the Block table for read
#         BlockTable acBlkTbl;
#         acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
#                                         OpenMode.ForRead) as BlockTable;

#         // Open the Block table record Model space for read
#         BlockTableRecord acBlkTblRec;
#         acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
#                                         OpenMode.ForRead) as BlockTableRecord;

#         object acObj = null;
#         foreach (ObjectId acObjId in acBlkTblRec)
#         {
#             // Get the first object in Model space
#             acObj = acTrans.GetObject(acObjId,
#                                         OpenMode.ForRead);

#             break;
#         }

#         // Open the DimStyle table for read
#         DimStyleTable acDimStyleTbl;
#         acDimStyleTbl = acTrans.GetObject(acCurDb.DimStyleTableId,
#                                             OpenMode.ForRead) as DimStyleTable;

#         string[] strDimStyleNames = new string[3];
#         strDimStyleNames[0] = "Style 1 copied from a dim";
#         strDimStyleNames[1] = "Style 2 copied from Style 1";
#         strDimStyleNames[2] = "Style 3 copied from the running drawing values";

#         int nCnt = 0;

#         // Keep a reference of the first dimension style for later
#         DimStyleTableRecord acDimStyleTblRec1 = null;

#         // Iterate the array of dimension style names
#         foreach (string strDimStyleName in strDimStyleNames)
#         {
#             DimStyleTableRecord acDimStyleTblRec;
#             DimStyleTableRecord acDimStyleTblRecCopy = null;

#             // Check to see if the dimension style exists or not
#             if (acDimStyleTbl.Has(strDimStyleName) == false)
#             {
#                 if (acDimStyleTbl.IsWriteEnabled == false) acDimStyleTbl.UpgradeOpen();

#                 acDimStyleTblRec = new DimStyleTableRecord();
#                 acDimStyleTblRec.Name = strDimStyleName;

#                 acDimStyleTbl.Add(acDimStyleTblRec);
#                 acTrans.AddNewlyCreatedDBObject(acDimStyleTblRec, true);
#             }
#             else
#             {
#                 acDimStyleTblRec = acTrans.GetObject(acDimStyleTbl[strDimStyleName],
#                                                         OpenMode.ForWrite) as DimStyleTableRecord;
#             }

#             // Determine how the new dimension style is populated
#             switch ((int)nCnt)
#             {
#                 // Assign the values of the dimension object to the new dimension style
#                 case 0:
#                     try
#                     {
#                         // Cast the object to a Dimension
#                         Dimension acDim = acObj as Dimension;

#                         // Copy the dimension style data from the dimension and
#                         // set the name of the dimension style as the copied settings
#                         // are unnamed.
#                         acDimStyleTblRecCopy = acDim.GetDimstyleData();
#                         acDimStyleTblRec1 = acDimStyleTblRec;
#                     }
#                     catch
#                     {
#                         // Object was not a dimension
#                     }

#                     break;

#                 // Assign the values of the dimension style to the new dimension style
#                 case 1:
#                     acDimStyleTblRecCopy = acDimStyleTblRec1;
#                     break;
#                 // Assign the values of the current drawing to the dimension style
#                 case 2:
#                     acDimStyleTblRecCopy = acCurDb.GetDimstyleData();
#                     break;
#             }

#             // Copy the dimension settings and set the name of the dimension style
#             acDimStyleTblRec.CopyFrom(acDimStyleTblRecCopy);
#             acDimStyleTblRec.Name = strDimStyleName;

#             // Dispose of the copied dimension style
#             acDimStyleTblRecCopy.Dispose();

#             nCnt = nCnt + 1;
#         }

#         // Commit the changes and dispose of the transaction
#         acTrans.Commit();
#     }
# }