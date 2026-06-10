  
import acad
import academit
import System


def 命令(): 
    academit.添加命令("rob", roa)
    academit.添加命令("roc", roc)
    academit.添加命令("xlv", xlv)
    academit.添加命令("xlh", xlh)
    academit.添加命令("zooma", zooma)
    academit.添加命令("dset", dset)
    academit.添加命令("ungroupdict", ungroupdict)
    
@acad.decorator_command
def roc():
    ss1 = acad.SSGet()
    pt1 = acad.GetPoint("选择基点:")
    acad.Command(["rotate", ss1, "", acad.ToPoint3d(pt1), "C"])
 
@acad.decorator_command
def xlv():
    acad.Command(["xline", "v"])

@acad.decorator_command
def xlh():
    acad.Command(["xline", "h"])


@acad.decorator_command
def zooma():
    acad.Command(["zoom", "a"])

@acad.decorator_command
def roa():
    while True:
        ss1 = acad.SSGet()
        if ss1 == None: return
        pt1, pt2 = acad.GetPoint2()
        if pt1 == None: return
        angle = acad.GetDouble(-90)
        acad.Command(["rotate3d", ss1, "", acad.ToPoint3d(pt1), acad.ToPoint3d(pt2), angle])


@acad.decorator_command
def dset():
    objid = acad.db.Dimstyle
    with acad.transaction() as trans:
        record = acad.TransObjectForRead(objid)
        dimscale = record.Dimscale
    numb = int(dimscale)
    biaohao = acad.GetInt(numb, "请输入标注号:")
    acad.SetCurrentDimStyle(f"副本{biaohao} ISO-25")

@acad.decorator_command
def ungroupdict():
    with acad.transaction() as trans:
        groupdict = acad.TransObjectForWrite(acad.db.GroupDictionaryId)
        for ide in groupdict: # DictionaryEntry 'DictionaryEntry' object has no attribute 'Erase'
            partgroup = acad.TransObjectForWrite(ide.Value) # Erase 必须配合 Write 对象使用
            partgroup.Erase(True)
            # partgroup = acad.TransObjectForRead(ide.Value)
            # partgroup.UpgradeOpen()
            # partgroup.Erase(True)
            # partgroup.DowngradeOpen()
