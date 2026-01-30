  
import acad
import academit
import System

def 命令(): 
    academit.添加命令("llblock", llblock)
    academit.添加命令("llexplode", llexplode)


@acad.decorator_command
def llblock():
    objidlist =  acad.SSGetIdList()   
    pt1, pt2 = acad.GetIdListBoundXY0(objidlist)
    # print(objidlist, pt1)
    with acad.transaction() as trans:
        acad.AddBlockFromIdList(objidlist, pt1)


@acad.decorator_command
def llexplode():
    objidlist =  acad.SSGetIdList()   
    with acad.transaction() as trans:
        for objid in objidlist:
            result_objidlist = acad.TransExplode(objid)
            # print(result_objidlist)

