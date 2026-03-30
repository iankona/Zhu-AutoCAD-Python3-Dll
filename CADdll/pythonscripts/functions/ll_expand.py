  
import acad
import academit
import System


def 命令(): 
    academit.添加命令("roc", roc)
    academit.添加命令("xlv", xlv)
    academit.添加命令("xlh", xlh)
    academit.添加命令("zooma", zooma)

    
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