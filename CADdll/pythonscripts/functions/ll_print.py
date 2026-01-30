import clr

import acad
import academit

import System

def 命令(): 
    academit.添加命令("ll-print", ll_print)
    academit.添加命令("ll-print-bound", ll_print_bound)    
    academit.添加命令("ll-print-corner", ll_print_corner)
    academit.添加命令("ll-print-corner-cross", ll_print_corner_cross)
    academit.添加命令("ll-print-frence", ll_print_frence)
    academit.添加命令("ll-print-entsel", ll_print_entsel)
    academit.添加命令("ll-print-clayer", ll_print_current_layer)
    academit.添加命令("ll-print-point-in-bound", ll_print_point_in_bound)
    academit.添加命令("ll-print-include", ll_print_include)
    academit.添加命令("ll-print-dr1dr2-angle", ll_print_dr1dr2_angle)


@acad.decorator_command
def ll_print_dr1dr2_angle():
    pt1 = acad.GetPoint("请选择dr1第1个顶点:")
    pt2 = acad.GetPoint("请选择dr1第2个顶点:")
    po1 = acad.GetPoint("请选择dr+2第1个顶点:")
    po2 = acad.GetPoint("请选择dr+2第2个顶点:")
    pd1 = acad.GetPoint("请选择dr-2第1个顶点:")
    pd2 = acad.GetPoint("请选择dr-2第2个顶点:")    
    dr1 = acad.Direct(pt1, pt2)
    dr2 = acad.Direct(po1, po2)
    dr3 = acad.Direct(pd1, pd2)
    angle1 = acad.AngleFromDotDr1Dr2(dr1, dr2)
    angle2 = acad.AngleFromDotDr1Dr2(dr1, dr3)
    angle3 = acad.AngleFromCrossDr1Dr2(dr1, dr2)
    angle4 = acad.AngleFromCrossDr1Dr2(dr1, dr3)
    dot1 = acad.Dot(dr1, dr2)
    dot2 = acad.Dot(dr1, dr3)
    cross1 = acad.Cross(dr1, dr2)
    cross2 = acad.Cross(dr1, dr3)
    acad.Prompt(f"dot: {dot1}, {dot2}\n")
    acad.Prompt(f"cross: {cross1}, {cross2}\n")
    acad.Prompt(f"dot angle: {angle1}, {angle2}\n")
    acad.Prompt(f"cross angle: {angle3}, {angle4}\n")
# dotangle: 18.48204843850921, 18.48204843850897
# crossangle: 18.482048438509214, 18.482048438508965
# dotangle: 161.5179515203917, 161.51795152039148
# crossangle: 18.482048438508965, 18.482048438509214
# dot: 749274.8775050562, 749274.8775050562
# cross: [0.0, -0.0, -250442.85412386296], [0.0, -0.0, 250442.85412385932]
# cross: [55683.07415465842, 33180.526255073746, -241909.22080778686], [-55683.074154658025, -33180.52625507285, 241909.2208077802]


@acad.decorator_command
def ll_print_include():
    objid1 = acad.EntSel()
    objid2 = acad.EntSel()
    acad.IsInclude(objid1, objid2)



@acad.decorator_command
def ll_print():
    objid = acad.EntSelSub()
    with acad.transaction() as trans:
        pass
        # objref = acad.TransObjectForRead(objid)
        # acad.Prompt(objref.GetType()) # Autodesk.AutoCAD.DatabaseServices.DBPoint


@acad.decorator_command
def ll_print_bound():
    objid = acad.EntSel()
    pt1, pt2 = acad.GetEntityBound(objid)
    acad.CommandAddPoint(pt1)
    acad.CommandAddPoint(pt2)



@acad.decorator_command
def ll_print_point_in_bound():
    objid = acad.EntSel()
    pt0 = acad.GetPoint()
    pt1, pt2 = acad.GetEntityBound(objid)
    acad.CommandAddPoint(pt0)
    acad.CommandAddPoint(pt1)
    acad.CommandAddPoint(pt2)
    flag  = acad.IsPointInRange(pt0, objid)
    acad.Prompt(flag)



@acad.decorator_command
def ll_print_entsel():
    ss1 =  acad.SSGet(sel_method=":S")
    acad.Prompt(ss1)



@acad.decorator_command
def ll_print_frence():
    pt1 = acad.GetPoint()
    pt2 = acad.GetPoint(base_point=pt1)
    # ss1 =  acad.SSGet(sel_method="+F")
    ss1 = acad.GetSelectFence(pt1, pt2)
    acad.Prompt(ss1)

@acad.decorator_command
def ll_print_corner_cross():
    pt1 = acad.GetPoint()
    pt2 = acad.GetCorner("", pt1)
    # ss1 =  acad.SSGet(sel_method="+F")
    ss1 = acad.GetSelectCornerCross(pt1, pt2)
    acad.Prompt(ss1)


@acad.decorator_command
def ll_print_corner():
    pt1 = acad.GetPoint()
    pt2 = acad.GetCorner("", pt1)
    # ss1 =  acad.SSGet(sel_method="+F")
    ss1 = acad.GetSelectCorner(pt1, pt2)
    acad.Prompt(ss1)



@acad.decorator_command
def ll_print_current_layer():
    with acad.transaction() as trans:
        objid = acad.db.Clayer
        layer_record = acad.GetObjectForRead(objid)
        acad.Prompt(layer_record.Name), acad.Prompt("\n")        
        acad.Prompt(layer_record.Color.ColorIndex)



# [CommandMethod("SelectNestedEntity")]
#         public static void SelectNestedEntity_Method()
#         {
#             Database db = HostApplicationServices.WorkingDatabase;
#             Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;

#             try
#             {
#                 PromptNestedEntityOptions nestedEntOpt = new PromptNestedEntityOptions("\nPick a nested entity:");
#                 PromptNestedEntityResult nestedEntRes = ed.GetNestedEntity(nestedEntOpt);
                

#                 if (nestedEntRes.Status == PromptStatus.OK)
#                 {
                               
#                     ObjectId[] ids = nestedEntRes.GetContainers();
#                     ObjectId target_id = nestedEntRes.ObjectId;
#                     ObjectId[] result_ids = new ObjectId[ids.Length + 1];
#                     for (int i = ids.Length - 1; i >= 0; i--)
#                         result_ids[(ids.Length - 1) - i] = ids[i];

#                     result_ids[ids.Length] = target_id;
#                     SubentityId subEnt = new SubentityId(SubentityType.Null, IntPtr.Zero);
#                     FullSubentityPath path = new FullSubentityPath(result_ids, subEnt);
#                     using (Transaction tr = db.TransactionManager.StartTransaction())
#                     {
#                         //Open the top nest
#                         Entity ent = (Entity)tr.GetObject(result_ids[0], OpenMode.ForRead);                      
#                         ent.Highlight(path, false);
#                         tr.Commit();
#                     }
#                 }
#             }
#             catch (System.Exception ex)
#             {
#                 ed.WriteMessage(ex.ToString());
#             }
#         }


# [CommandMethod("SEL")
# public void Selection()
# {
#     var document = Application.DocumentManager!.MdiActiveDocument!;
#     var editor = document.Editor!;
#     var result = editor.GetSelection()!;
#     if (result.Status == PromptStatus.Cancel) return;
#     var solid3dClass = RXObject.GetClass(typeof(Solid3d));
#     using var transaction = document.TransactionManager!.StartTransaction()!;
#     foreach (SelectedObject selectedObject in result.Value!)
#     {
#         if (selectedObject.ObjectId.ObjectClass == solid3dClass)
#         {
#             var solid3d = (Solid3d)selectedObject.ObjectId.GetObject(OpenMode.ForRead)!;
#             if (selectedObject is PickPointSelectedObject pickPointSelectedObject)
#             {
#                 FullSubentityPath[]? paths = solid3d.GetSubentityPathsAtGraphicsMarker(SubentityType.Edge, selectedObject.GraphicsSystemMarkerPtr, pickPointSelectedObject.PickPoint.PointOnLine, Matrix3d.Identity, null!);
#                 if (paths is { Length: > 0 }) solid3d.Highlight(paths[0], highlightAll: false);
#             }
#         }
#         else
#         {
#             Debug.Print("Selected object: " + selectedObject.SelectionMethod);
#         }
#     }
#     transaction!.Commit();
# }