  
import acad
import academit
import System


def 命令(): 
    academit.添加命令("roc", roc)
    academit.添加命令("xlv", xlv)
    academit.添加命令("xlh", xlh)

@acad.decorator_command
def roc():
    ss1 = acad.SSGet()
    pt1 = acad.GetPoint("选择基点:")
    # angle1 = acad.GetDouble(0, "旋转角度:")
    # acad.CommandRotateCopy(ss1, pt1, angle1)
    acad.Command(["rotate", ss1, "", acad.ToPoint3d(pt1), "C"])
 
@acad.decorator_command
def xlv():
    acad.Command(["xline", "v"])

@acad.decorator_command
def xlh():
    acad.Command(["xline", "h"])