  
import acad
import academit

def 命令(): 
    academit.添加命令("llbo-for", llbo_for)
    # academit.添加命令("ll-entsel", ll_entsel)


count = 1
@acad.decorator_command
def llbo_for():
    global count
    if count > 255: count = 1

    # zoom
    # ss1 = acad.SSGet()
    # pt1, pt2 = acad.GetBound(ss1)
    # with acad.command_undo(), acad.command_osmode():
    #      acad.CommandZoom(pt1, pt2)
         
    pt1 = acad.GetPoint("Select internal point: ")
    objreflist = acad.ed.TraceBoundary(acad.ToPoint3d(pt1), True)
    with acad.transaction() as trans:
        for i, objref in enumerate(objreflist):
              objref.ColorIndex = count 
              acad.AddDBObject(objref)
    count +=1 

 