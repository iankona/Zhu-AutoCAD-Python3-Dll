import clr

import acad
import academit

import System

# from Autodesk.AutoCAD.Interop.Common import AcPreviewMode, AcPlotScale, AcPlotType
# from System.Drawing import Image
# from Autodesk.AutoCAD.GraphicsSystem import View, KernelDescriptor, Manager, Device, RendererType
# from Autodesk.AutoCAD.UniqueString import Intern


def 命令(): 
    academit.添加命令("jpg1", jpg1)
    academit.添加命令("jpg2", jpg2)
    academit.添加命令("jpg3", jpg3)
    academit.添加命令("jpgx", jpgx)
    # academit.添加命令("jpgblock", jpgblock)



@acad.decorator_command
def jpg1():
    bmp = acad.DocumentExtension.CapturePreviewImage(acad.doc, System.UInt32(4096), System.UInt32(2160))
    System.Windows.Forms.Clipboard.SetDataObject(bmp) # 剪切板

@acad.decorator_command
def jpg2():
    bmp = acad.DocumentExtension.CapturePreviewImage(acad.doc, System.UInt32(7680), System.UInt32(4320)) # Bitmap
    System.Windows.Forms.Clipboard.SetDataObject(bmp) # 剪切板

@acad.decorator_command
def jpg3():
    bmp = acad.DocumentExtension.CapturePreviewImage(acad.doc, System.UInt32(15360), System.UInt32(8640))
    System.Windows.Forms.Clipboard.SetDataObject(bmp) # 剪切板

@acad.decorator_command
def jpgx():
    ratio = acad.GetDouble(2.0, "请输入超分辨率系数:")
    a = int(1920*ratio)
    b = int(1080*ratio)
    bmp = acad.DocumentExtension.CapturePreviewImage(acad.doc, System.UInt32(a), System.UInt32(b))
    System.Windows.Forms.Clipboard.SetDataObject(bmp) # 剪切板




# @acad.decorator_command
# def jpg00():
#     # with acad.transaction() as trans:
#     bmp = acad.db.ThumbnailBitmap
#     System.Windows.Forms.Clipboard.SetDataObject(bmp) # 剪切板


# @acad.decorator_command
# def jpg01():
#     view = acad.ed.GetCurrentView()
#     bmp = view.Thumbnail
#     System.Windows.Forms.Clipboard.SetDataObject(bmp) 





# Autodesk.AutoCAD.GraphicsSystem View

@acad.decorator_command
def jpgblock():
    flist = [
        [-4, "<OR"],[0, "*INSERT"], [2, "BLOCKDEFAULT*"], [-4, "OR>"]
        ]
    objid = acad.EntSel(flist)
    ptr = acad.Utils.GetBlockImage(objid, 1024, 1024, acad.Color.FromRgb(System.Byte(0),System.Byte(0),System.Byte(0)))
    bmp = Image.FromHbitmap(System.IntPtr(ptr)) # GDI+ 一般性错误
    System.Windows.Forms.Clipboard.SetDataObject(bmp) 



@acad.decorator_command
def jpgview():
    pass
#     gsm = acad.doc.GraphicsManager
#     view = View()
#     cvport = acad.Application.GetSystemVariable("CVPORT")
#     acad.doc.GraphicsManager.SetViewFromViewport(view, cvport)
#     view.SetView(acad.Point3d(0, 0, 1), acad.Point3d.Origin, acad.Vector3d.YAxis, 2048, 1096)
#     descriptor = KernelDescriptor()
#     descriptor.addRequirement(Intern("3D Drawing"))
#     kernel = Manager.AcquireGraphicsKernel(descriptor)
#     dev = gsm.CreateAutoCADOffScreenDevice(kernel)
#     dev.OnSize(pictureBox.Size);
#     dev.DeviceRenderType = RendererType.Default
#     dev.BackgroundColor = Color.Black
#     dev.Add(view);
#     dev.Update();
#     try
#     {
#       using (Model model = gsm.CreateAutoCADModel(kernel))
#       {
#         var extents3D = new Extents3d();
#         var line = new Line(p1, p2);
        
#         [..]
        
#         view.Add(line, model);
#         extents3D.AddExtents(line.Bounds.Value);
        
#         view.ZoomExtents(extents3D.MinPoint, extents3D.MaxPoint);
#         view.Zoom(0.9);
#       }
#       pictureBox.Image = view.GetSnapshot(new Rectangle(0, 0, pictureBox.Width - 1,
#         pictureBox.Height - 1));
#     finally
#     {
#       view.EraseAll();
#       dev.Erase(view);
#       if (null != line)
#         line.Dispose();      
#     }
#   }
# }






# using Autodesk.AutoCAD.ApplicationServices;

# using Autodesk.AutoCAD.DatabaseServices;

# using Autodesk.AutoCAD.EditorInput;

# using Autodesk.AutoCAD.GraphicsInterface;

# using Autodesk.AutoCAD.GraphicsSystem;

# using Autodesk.AutoCAD.Runtime;

# using Autodesk.AutoCAD.Interop;

# using System.Drawing;

# namespace OffscreenImageCreation

# {

#   public class Commands

#   {

#     [CommandMethod("OSS")]

#     static public void OffscreenSnapshot()

#     {

#       CreateSphere();

#       SnapshotToFile(

#         "c:\\sphere-Wireframe2D.png",

#         VisualStyleType.Wireframe2D

#       );

#       SnapshotToFile(

#         "c:\\sphere-Hidden.png",

#         VisualStyleType.Hidden

#       );

#       SnapshotToFile(

#         "c:\\sphere-Basic.png",

#         VisualStyleType.Basic

#       );

#       SnapshotToFile(

#         "c:\\sphere-ColorChange.png",

#         VisualStyleType.ColorChange

#       );

#       SnapshotToFile(

#         "c:\\sphere-Conceptual.png",

#         VisualStyleType.Conceptual

#       );

#       SnapshotToFile(

#         "c:\\sphere-Flat.png",

#         VisualStyleType.Flat

#       );

#       SnapshotToFile(

#         "c:\\sphere-Gouraud.png",

#         VisualStyleType.Gouraud

#       );

#       SnapshotToFile(

#         "c:\\sphere-Realistic.png",

#         VisualStyleType.Realistic

#       );

#     }

#     static public void CreateSphere()

#     {

#       Document doc =

#         Application.DocumentManager.MdiActiveDocument;

#       Database db = doc.Database;

#       Editor ed = doc.Editor;

#       Transaction tr =

#         doc.TransactionManager.StartTransaction();

#       using (tr)

#       {

#         BlockTable bt =

#           (BlockTable)tr.GetObject(

#             db.BlockTableId,

#             OpenMode.ForRead

#           );

#         BlockTableRecord btr =

#           (BlockTableRecord)tr.GetObject(

#             bt[BlockTableRecord.ModelSpace],

#             OpenMode.ForWrite

#           );

#         Solid3d sol = new Solid3d();

#         sol.CreateSphere(10.0);

#         const string matname =

#           "Sitework.Paving - Surfacing.Riverstone.Mortared";

#         DBDictionary matdict =

#           (DBDictionary)tr.GetObject(

#             db.MaterialDictionaryId,

#             OpenMode.ForRead

#           );

#         if (matdict.Contains(matname))

#         {

#           sol.Material = matname;

#         }

#         else

#         {

#           ed.WriteMessage(

#             "\nMaterial (" + matname + ") not found" +

#             " - sphere will be rendered without it.",

#             matname

#           );

#         }

#         btr.AppendEntity(sol);

#         tr.AddNewlyCreatedDBObject(sol, true);

#         tr.Commit();

#       }

#       AcadApplication acadApp =

#         (AcadApplication)Application.AcadApplication;

#       acadApp.ZoomExtents();

#     }

#     static public void SnapshotToFile(

#       string filename,

#       VisualStyleType vst

#     )

#     {

#       Document doc =

#         Application.DocumentManager.MdiActiveDocument;

#       Editor ed = doc.Editor;

#       Database db = doc.Database;

#       Manager gsm = doc.GraphicsManager;

#       // Get some AutoCAD system variables

#       int vpn =

#         System.Convert.ToInt32(

#           Application.GetSystemVariable("CVPORT")

#         );

#       // Get AutoCAD's GS view for this document...

#       View gsv =

#         doc.GraphicsManager.GetGsView(vpn, true);

#       // ... but create a new one for the actual snapshot

#       using (View view = new View())

#       {

#         // Set the view to be just like the one

#         // in the AutoCAD editor

#         view.Viewport = gsv.Viewport;

#         view.SetView(

#           gsv.Position,

#           gsv.Target,

#           gsv.UpVector,

#           gsv.FieldWidth,

#           gsv.FieldHeight

#         );

#         // Set the visual style to the one passed in

#         view.VisualStyle = new VisualStyle(vst);

#         Device dev =

#           gsm.CreateAutoCADOffScreenDevice();

#         using (dev)

#         {

#           dev.OnSize(gsm.DisplaySize);

#           // Set the render type and the background color

#           dev.DeviceRenderType = RendererType.Default;

#           dev.BackgroundColor = Color.White;

#           // Add the view to the device and update it

#           dev.Add(view);

#           dev.Update();

#           using (Model model = gsm.CreateAutoCADModel())

#           {

#             Transaction tr =

#               db.TransactionManager.StartTransaction();

#             using (tr)

#             {

#               // Add the modelspace to the view

#               // It's a container but also a drawable

#               BlockTable bt =

#                 (BlockTable)tr.GetObject(

#                   db.BlockTableId,

#                   OpenMode.ForRead

#                 );

#               BlockTableRecord btr =

#                 (BlockTableRecord)tr.GetObject(

#                   bt[BlockTableRecord.ModelSpace],

#                   OpenMode.ForRead

#                 );

#               view.Add(btr, model);

#               tr.Commit();

#             }

#             // Take the snapshot

#             Rectangle rect = view.Viewport;

#             using (Bitmap bitmap = view.GetSnapshot(rect))

#             {

#               bitmap.Save(filename);

#               ed.WriteMessage(

#                 "\nSnapshot image saved to: " +

#                 filename

#               );

#               // Clean up

#               view.EraseAll();

#               dev.Erase(view);

#             }

#           }

#         }

#       }

#     }

#   }

# }


# Autodesk.AutoCAD.Windows.Data.CMLContentSearchPreviews 类（acmgd.dll中），
# 它有一个静态方法GetBlockTRThumbnail（BlockTableRecord）。
# 这个方法是为AutoCAD设计的，可以在运行时获取一个块的缩略图，用于AutoCAD的UI，所以可能正是你想要的。
# 好处是它直接生成缩略图作为ImageSource，所以你不需要代码把位图转换成ImageSource用于AutoCAD UI

# ==============================================================

# UCS需要先转WCS，然后从WCS转DCS

# 第一步UCS TO WCS

# point.TransformBy(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.CurrentUserCoordinateSystem);

# 妥妥的

# 然后第二部复杂一点：

# // 将 WCS 坐标变换为 DCS 坐标     

# Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument; 

# ViewTableRecord acView = acDoc.Editor.GetCurrentView();

# Matrix3d matWCS2DCS;           

# matWCS2DCS = Matrix3d.PlaneToWorld(acView.ViewDirection);           

# matWCS2DCS = Matrix3d.Displacement(acView.Target - Point3d.Origin) * matWCS2DCS;           

# matWCS2DCS = Matrix3d.Rotation(-acView.ViewTwist, acView.ViewDirection, acView.Target) * matWCS2DCS;

# Extents3d extent = new Extents3d(pt1, pt2);

# matWCS2DCS = matWCS2DCS.Inverse();

# extent.TransformBy(matWCS2DCS);

# ==============================================================













@acad.decorator_command
def png1():
    isZong = True
    isZong = PrintPlotRotation()

    acad.doc.SetVariable("BACKGROUNDPLOT", 0)
    acad.doc.Application.ZoomExtents()
    acad.doc.ActiveLayout.ConfigName = "PublishToWeb JPG.pc3"
    cMNameLst = acad.doc.ActiveLayout.GetCanonicalMediaNames()

    acad.doc.ActiveLayout.CenterPlot = True # 居中
    acad.doc.ActiveLayout.PlotType = AcPlotType.acExtents
    acad.doc.ActiveLayout.UseStandardScale = True # 使用标准比例
    acad.doc.ActiveLayout.StandardScale = AcPlotScale.acScaleToFit # 自动缩放适应

    # // 设置打印样式
    # doc.ActiveLayout.StyleSheet = "acad.ctb"
    # if (isZong)
    # {
    #     doc.ActiveLayout.PlotRotation = AcPlotRotation.ac0degrees; //纵向打印 
    # }
    # else
    #     doc.ActiveLayout.PlotRotation = AcPlotRotation.ac270degrees; //纵向打印 
    # //打印预览
    acad.doc.Plot.DisplayPlotPreview(AcPreviewMode.acFullPreview)
    acad.doc.Plot.QuietErrorMode = True # 生成存档，避免报错
    acad.doc.Plot.NumberOfCopies = 1 # 打印份数
    acad.doc.ActiveLayout.RefreshPlotDeviceInfo()
    acad.doc.Plot.PlotToFile("F:\\U桌面备份\\Documents\\", "PublishToWeb JPG.pc3") # 打印到文件，第二个参数为打印机名称


    # foreach (string name in cMNameLst)
    # {
    #     //查找纸张大小
    #     if (name.Contains("2000.00") && name.Contains("2000.00"))
    #     {
    #         drawing.ActiveLayout.CanonicalMediaName = name;
    #         break;
    #     }
    # }

# =======================================================================
# 做个标记备忘。// Helper function to generate an Image from a BitmapSource
#         // 函数：从位图源生成190*120图像
#         private static Image ImageSourceToGDI(BitmapSource src)
#         {
#             var ms = new MemoryStream();
#             var encoder = new BmpBitmapEncoder();
#             encoder.Frames.Add(BitmapFrame.Create(src));
#             encoder.Save(ms);
#             ms.Flush();
#             return Image.FromStream(ms);
#         }
# 1.界面
# private void BlockPreviews2_Load(object sender, EventArgs e)
#         {
#             Database db = HostApplicationServices.WorkingDatabase;
#             using (Transaction trans = db.TransactionManager.StartTransaction())
#             {
#                 //打开块表
#                 var bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
#                 var namelist = new List<string>();//新建图块名称表
#                 ImageList imglist = new ImageList();//新建图块缩略图表
#                 imglist.ImageSize = new Size(190,120);// 图片尺寸分别是宽和高
#                 foreach (ObjectId blockRecordId in bt)//循环遍历块表中的块表记录
#                 {
#                     var btr = (BlockTableRecord)trans.GetObject(blockRecordId, OpenMode.ForRead);//打开块表记录对象
#                     if (!btr.IsAnonymous && !btr.IsLayout)//在下拉列表框中只加入非匿名块和非布局块的名称
#                     {
#                         namelist.Add(btr.Name);//添加到图块名称表
#                         imglist.Images.Add(ImageSourceToGDI(CMLContentSearchPreviews.GetBlockTRThumbnail(btr) as BitmapSource));//添加到图块缩略图表
#                     }
#                 }
#                 listView1.View = View.LargeIcon;//设置为大图标视图
#                 listView1.MultiSelect = false; //只能单选
#                 listView1.LargeImageList = imglist;// 这里设置listView的SmallImageList ,用imgList将其撑大
#                 for (int i = 0; i < imglist.Images.Count; i++)
#                 {
#                     var lvi = new ListViewItem();//新建ListViewItem
#                     lvi.ImageIndex = i;//取出图片
#                     lvi.Text = namelist[i];//取出图块名称
#                     listView1.Items.Add(lvi);//添加到listView1
#                 }
#             }
#         }


# 2.获取选中项
# private void listView1_SelectedIndexChanged(object sender, EventArgs e)
#         {
#             if (listView1.FocusedItem != null) //这个if必须的，不然会得到值但会报错
#             {
#                 textBox1.Text = "当前选中的是:" + listView1.FocusedItem.SubItems[0].Text;//获得的listView的值显示在文本框里
#             }
#         }






# imf = System.Drawing.Imaging.ImageFormat.Bmp
# imf = System.Drawing.Imaging.ImageFormat.Gif
# imf = System.Drawing.Imaging.ImageFormat.Jpeg
# imf = System.Drawing.Imaging.ImageFormat.Tiff
# imf = System.Drawing.Imaging.ImageFormat.Wmf
# imf = System.Drawing.Imaging.ImageFormat.Png
# pofo = acad.PromptSaveFileOptions
# #pofo.Filter = "Bitmap (*.bmp)|*.bmp|GIF (*.gif)|*.gif|JPEG (*.jpg)|*.jpg|PNG (*.png)|*.png|TIFF (*.tif)|*.tif"
# pfnr = acad.ed.GetFileNameForSave("鹅鹅鹅")
# # if pfnr.Status != acad.PromptStatus.OK: return
# outFile = pfnr.StringResult
# bmp.Save(outFile, imf)



# def ll_entsel():
#     acad.GetActiveDocument()
#     acad.EntSel()



# object oCad = Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication;
# Type tpCad = oCad.GetType();
# object oDoc = tpCad.InvokeMember("ActiveDocument", System.Reflection.BindingFlags.GetProperty, null, oCad, null);
# Type tpDoc = oDoc.GetType();
# object ass = tpDoc.InvokeMember("ActiveSelectionSet", System.Reflection.BindingFlags.GetProperty, null, oDoc, null);
# tpDoc.InvokeMember("Export", System.Reflection.BindingFlags.InvokeMethod, null, oDoc, new object[] { path + "\\abc", "WMF", ass });
# ed.SetImpliedSelection(new ObjectId[] { });


