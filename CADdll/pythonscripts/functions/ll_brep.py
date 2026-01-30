  
import acad
import academit

def 命令(): 
    academit.添加命令("ll-brep", ll_brep)
    academit.添加命令("ll-select-solid-edge", ll_select_solid_edge)


@acad.decorator_command
def ll_brep():
    objid = acad.EntSel([[0, "REGION"]])
    with acad.transaction() as trans:
        region = acad.TransObjectForWrite(objid)
    brep = acad.Brep(region)
    edges = brep.Edges
    for edge in edges:
        curve = edge.Curve
        print("样条曲线、原始曲线、线段、线、弧", curve.IsNurbCurve, curve.IsNativeCurve, curve.IsLineSegment, curve.IsLine, curve.IsCircularArc)    # 样条曲线、原始曲线、线段、线、弧 False True True False False                  
                         


@acad.decorator_command
def ll_select_solid_edge():
    viewrecode = acad.ed.GetCurrentView() # ViewTableRecord
    print(viewrecode)
    objid = viewrecode.Layout
    print(objid)
    with acad.transaction() as trans:
        objref = acad.TransObjectForWrite(objid)
    acad.Prompt(objref.GetType())


# opts = new PromptSelectionOptions();
# opts.AllowSubSelections = true;
# AllowSubSelections allows to select sub entities (solid or meshes edges, attribute references, ...) accordingly with the LEGACYCTRLPICK and SUBOBJSELECTMODE sysvar settings.
# With default settings (LEGACYCTRLPICK = 2 and SUBOBJSELECTMODE = 0), AllowSubSelections allows sub entities selection using Ctrl+click.
# 你好
# AllowSubSelections 允许通过 LEGACYCTRLPICK 和 SUBOBJSELECTMODE sysvar 设置，相应地选择子实体（实体或网格、边缘、属性引用等）。
# 在默认设置（LEGACYCTRLPICK = 2 和 SUBOBJSELECTMODE = 0）下，允许子实体选择时，可以通过 Ctrl+click 选择子实体。