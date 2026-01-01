import clr

import acad
import academit
from Autodesk.AutoCAD.Geometry import Point2d, Point3d, Point3dCollection, Matrix3d, Vector3d
import System

def 命令(): 
    academit.添加命令("llregion-print", llregion_print)

    pass


@acad.decorator_command
def llregion_print():
    objidlist = acad.SSGetIdList()
    with acad.transaction() as trans:
        for objid in objidlist:
            objref = acad.GetObjectForWrite(objid)
            normal = objref.Normal # (0.554700196225229,0,0.832050294337844) 
            x,y,z = objref.Normal
            matrix4x41 = objref.Ecs # ((1,0,0,0),(0,1,0,0),(0,0,1,0),(0,0,0,1))
            matrix4x42 = Matrix3d.PlaneToWorld(Vector3d(0,0,1))
            matrix4x43 = Matrix3d.PlaneToWorld(objref.Normal)
            matrix4x44 = Matrix3d.PlaneToWorld(Vector3d(-x,0,1-z))
            print(objref, objref.Normal)
            objref.TransformBy(matrix4x44)
            print(objref, objref.Normal)

            