import clr
import random
import acad
import academit
from Autodesk.AutoCAD.DatabaseServices import SubentityId, SubentityType, FullSubentityPath, MeshFaceterData, IdMapping, ObjectIdCollection, SubDMesh, Curve, Extents3d, Polyline, Polyline3d, Line, Circle, Poly3dType, DBText, Region, BooleanOperationType
from Autodesk.AutoCAD.Geometry import Point2d, Point3d, Point3dCollection, Matrix3d, Vector3d, NurbCurve3d
import System
from Autodesk.AutoCAD.BoundaryRepresentation import PointContainment, Brep, Face, BrepEntity
import Autodesk
from Autodesk.AutoCAD.Colors import Color, ColorMethod

def 命令(): 
    academit.添加命令("llregion", llregion)
    academit.添加命令("llregion-cut", llregion_cut)
    academit.添加命令("llregion-subtract", llregion_subtract)
    academit.添加命令("llregion-print", llregion_print)
    academit.添加命令("llregion-check-inside", llregion_check_inside)
    academit.添加命令("llregion-mesh-sphm", llregion_mesh_sphm)
    pass


# 可以通过检查ed.CurrentUserCoordinateSystem是否为单位矩阵(Matrix3d.Identity) 来判断
# 如果ed.CurrentUserCoordinateSystem == Matrix3d.Identity, 则UCS与WCS一致, 无需转换
# Matrix.Inverse()



@acad.decorator_command
def llregion_mesh_sphm():
    with acad.transaction() as trans:
        pass
        # SphericalMesh
        mySubD = SubDMesh()
        mySubD.SetSphere(20, 8, 8, 0)
        mySubD.SetDatabaseDefaults()
        acad.AddDBObject(mySubD)

        # cloneAndInflateSolid
        solidId = mySubD.ObjectId
        frac = 0.2
        ids = ObjectIdCollection()
        ids.Add(solidId)
        im = IdMapping()
        acad.db.DeepCloneObjects(ids, acad.db.CurrentSpaceId, im, False)
        newId = im.Lookup(solidId).Value
        sol = acad.GetObjectForWrite(newId)
        ext = sol.Bounds
        vec = ext.MaxPoint - ext.MinPoint
        # # sol.OffsetBody(vec.Length * frac)
        # fd = MeshFaceterData(0.01 * vec.Length, 40 * System.Math.PI / 180, 2, 2, 15, 5, 5, 0 )
        # mdc = SubDMesh.GetObjectMesh(sol, fd)
        # sd = SubDMesh()
        # sd.SetDatabaseDefaults()
        # sd.SetSubDMesh(mdc.VertexArray, mdc.FaceArray, 1)
        # # sol.Erase()

        # meshId = newId
        # creasePaths =System.List[FullSubentityPath]
        # sd.UpgradeOpen()
        # subEnts = sd.GetSubentities()
        # for subEnt in subEnts:
        #     if subEnt.FullSubentityPath.SubentId.Type == SubentityType.Edge:
        #         creasePaths.Add(subEnt.FullSubentityPath)
 
        # sd.SetCrease(creasePaths.ToArray(), -1)

        # bumpPaths =System.List[FullSubentityPath]
        # sd.UpgradeOpen()
        # subEnts = sd.GetSubentities()

        # extrudelen = 1.0
        # rnd = System.Random()
 
        # for subEnt in subEnts:
        #     if subEnt.FullSubentityPath.SubentId.Type == SubentityType.Face:
        #         faces = FullSubentityPath[1]
        #         faces[0] = subEnt.FullSubentityPath
        #         fPlane = sd.GetFacePlane(faces[0].SubentId)
        #         norm = fPlane.Normal
        #         ext = sd.Bounds
        #         vec = ext.MaxPoint - ext.MinPoint
        #         mat = Matrix3d.Displacement(norm * 0.5 * extrudelen *  rnd.NextDouble() * vec.Length)
        #         sd.TransformSubentityPathsBy(faces, mat)

        # rnd = System.Random()
        # sd.UpgradeOpen()
        # for i in range(sd.NumberOfFaces - 1):
        #     sId = SubentityId(SubentityType.Face, i)
        #     r,g,b = random.randint(1, 255),random.randint(1, 255),random.randint(1, 255)
        #     col = Color.FromRgb(r,g,b)
        #     sd.SetSubentColor(sId, col)

        # sd.UpgradeOpen()
        # sol = sd.ConvertToSolid(True, True)
        # newId = acad.AddDBObject(sol)

 
        # ext = sol.Bounds
        # pt1 = Point3d(ext.MaxPoint.X, ext.MinPoint.Y, ext.MinPoint.Z)
        # transVec = 1.5 * (pt1 - ext.MinPoint)
        # mat = Matrix3d.Displacement(transVec)
        # sol.TransformBy(mat)

@acad.decorator_command
def llregion():
    dbobj_collect = acad.SSGetDbObjectCollection()
    with acad.transaction() as trans:
        regions = acad.AddRegion(dbobj_collect, "面域1", 35)
        area_dbobj_list = []
        for objref in regions: area_dbobj_list.append([objref.Area, objref])
        area_dbobj_list.sort(key = lambda item: item[0], reverse=True) # python调用了C#的排序出错，使用lambda使运行正常，原理不明 # Cannot convert object of type Python.Runtime.CLRObject to IComparable
        check_index_list = [0]
        for i, [area, objref] in enumerate(area_dbobj_list):
            if i in check_index_list: continue
            check_index_list.append(i)
            large_objref = objref
            child_objref_list = []
            for j, [area, objref] in enumerate(area_dbobj_list[i+1:]):
                extend = objref.GeometricExtents
                center = Point3d( (extend.MinPoint.X+extend.MaxPoint.X)/2, (extend.MinPoint.Y+extend.MaxPoint.Y)/2, (extend.MinPoint.Z+extend.MaxPoint.Z)/2 )
                brep = Brep(large_objref)
                result = brep.GetPointContainment(center, PointContainment.Outside) # out ref 被自动处理成结果返回
                brep_entity = result[0]
                if brep_entity == None: continue
                if str(brep_entity.GetType()) == "Autodesk.AutoCAD.BoundaryRepresentation.Face": # 点在面内
                    check_index_list.append(i+1+j)
                    child_objref_list.append(objref)
            for objref in child_objref_list: large_objref.BooleanOperation(BooleanOperationType.BoolSubtract, objref)
        area_dbobj_list[0][1].Erase(True)

# Brep.Edge to DB_Curve
# Brep brep = new Brep(pline);//出现异常
# BrepEdgeCollection elp = brep.Edges;
# foreach (Edge edge in elp)
# {
#     NurbCurve3d c3d = edge.GetCurveAsNurb();
#     Curve cv = Curve.CreateFromGeCurve(c3d);
# }
        
@acad.decorator_command
def llregion_sheet_cut():
    dbobj_collect = acad.SSGetDbObjectCollection()
    with acad.transaction() as trans:
        regions = acad.AddRegion(dbobj_collect, "面域1", 35)
        area_dbobj_list = []
        for objref in regions: area_dbobj_list.append([objref.Area, objref])
        area_dbobj_list.sort(key = lambda item: item[0], reverse=True) # 数值差距大，导致python调用了C#的排序出错，使用lambda使运行正常，原理不明
        check_index_list = [0]
        for i, [area, objref] in enumerate(area_dbobj_list):
            if i in check_index_list: continue
            check_index_list.append(i)
            large_objref = objref
            child_objref_list = []
            for j, [area, objref] in enumerate(area_dbobj_list[i+1:]):
                extend = objref.GeometricExtents
                center = Point3d( (extend.MinPoint.X+extend.MaxPoint.X)/2, (extend.MinPoint.Y+extend.MaxPoint.Y)/2, (extend.MinPoint.Z+extend.MaxPoint.Z)/2 )
                brep = Brep(large_objref)
                result = brep.GetPointContainment(center, PointContainment.Outside) # out ref 被自动处理成结果返回
                brep_entity = result[0]
                if brep_entity == None: continue
                if str(brep_entity.GetType()) == "Autodesk.AutoCAD.BoundaryRepresentation.Face": # 点在面内
                    check_index_list.append(i+1+j)
                    child_objref_list.append(objref)
            for objref in child_objref_list: large_objref.BooleanOperation(BooleanOperationType.BoolSubtract, objref)
        area_dbobj_list[0][1].Erase(True)
# (<Autodesk.AutoCAD.BoundaryRepresentation.Face object at 0x00000227DE697380>, <PointContainment.OnBoundary: 2>) Autodesk.AutoCAD.BoundaryRepresentation.Face
# (None, <PointContainment.Outside: 1>)




@acad.decorator_command
def llregion_cut():
    dbobj_collect = acad.SSGetDbObjectCollection()
    with acad.transaction() as trans:
        regions = acad.AddRegion(dbobj_collect, "面域1", 35)
        max_area, max_objref = 0, None
        dbobj_list = []
        for objref in regions:
        # for objid in objidlist:
        #     objref = acad.GetObjectForWrite(objid)
            dbobj_list.append(objref)
            if objref.Area > max_area:
                max_area = objref.Area
                max_objref = objref
        # print(max_area, max_objref)
        if max_objref != None:
            # dbobj_list.remove(max_objref)
            index = dbobj_list.index(max_objref)
            obj_list = dbobj_list[0:index] + dbobj_list[index+1:]
            for objref in obj_list: max_objref.BooleanOperation(BooleanOperationType.BoolSubtract, objref)


@acad.decorator_command
def llregion_check_inside():
    objid1 = acad.EntSel("请选择第1个对象:")
    objid2 = acad.EntSel("请选择第2个对象:")
    with acad.transaction() as trans:
        region1 = acad.GetObjectForWrite(objid1)
        region2 = acad.GetObjectForWrite(objid2)
        extend = region2.GeometricExtents
        point1 = extend.MinPoint
        point2 = extend.MaxPoint # [point1.X, point1.Y, point1.Z], [point2.X, point2.Y, point2.Z]         
        center_point = Point3d((point1.X+point2.X)/2, (point1.Y+point2.Y)/2, (point1.Z+point2.Z)/2)
        brep = Brep(region1)
        result = brep.GetPointContainment(center_point, PointContainment.Outside) # out ref 被自动处理成结果返回
        brep_entity = result[0]
        # print(brep_entity.GetType())
        # print(ent) # (<Autodesk.AutoCAD.BoundaryRepresentation.Face object at 0x00000204409347C0>, <PointContainment.OnBoundary: 2>)
        if str(brep_entity.GetType()) == "Autodesk.AutoCAD.BoundaryRepresentation.Face":
            acad.Prompt("点在面范围里")
            # region1.BooleanOperation(BooleanOperationType.BoolSubtract, region2)
            
@acad.decorator_command
def llregion_subtract():
    objidlist = acad.SSGetIdList([[0, "REGION"]])
    with acad.transaction() as trans:
        max_area, max_objref = 0, None
        dbobj_list = []
        for objid in objidlist:
            objref = acad.GetObjectForWrite(objid)
            dbobj_list.append(objref)
            if objref.Area > max_area:
                max_area = objref.Area
                max_objref = objref
        # print(max_area, max_objref)
        if max_objref != None:
            # dbobj_list.remove(max_objref)
            index = dbobj_list.index(max_objref)
            obj_list = dbobj_list[0:index] + dbobj_list[index+1:]
            for objref in obj_list: max_objref.BooleanOperation(BooleanOperationType.BoolSubtract, objref)



@acad.decorator_command
def llregion_print():
    objidlist = acad.SSGetIdList()
    with acad.transaction() as trans:
        for objid in objidlist:
            objref = acad.GetObjectForWrite(objid)
            normal = objref.Normal # (0.554700196225229,0,0.832050294337844) 
            x,y,z = objref.Normal
            matrix4x41 = objref.Ecs # ((1,0,0,0),(0,1,0,0),(0,0,1,0),(0,0,0,1))
            # matrix4x42 = Matrix3d.PlaneToWorld(Vector3d(0,0,1))
            # matrix4x43 = Matrix3d.PlaneToWorld(objref.Normal)
            # matrix4x44 = Matrix3d.PlaneToWorld(Vector3d(-x,0,1-z))
            # print(objref, objref.Normal)
            # objref.TransformBy(matrix4x44)
            # print(objref, objref.Normal)



# using Autodesk.AutoCAD.Geometry;

# Point3d pointA = new Point3d(1, 2, 3);
# Point3d pointB = new Point3d(4, 5, 6);
# Point3d pointC = new Point3d(7, 8, 9);

# // 计算向量
# Vector3d vectorAB = pointB - pointA;
# Vector3d vectorAC = pointC - pointA;

# // 点乘
# double dotProduct = vectorAB.DotProduct(vectorAC);

# // 叉乘
# Vector3d crossProduct = vectorAB.CrossProduct(vectorAC);

# // 输出结果
# Console.WriteLine($"点乘结果: {dotProduct}");
# Console.WriteLine($"叉乘结果: {crossProduct}");


import math
import numpy as np
origin_vector = np.array([1, 0 ,0])
location_vector = np.array([0, 0 ,1])
#注意，如果向量没有归一化，可以先考虑归一化下。
c = np.dot(origin_vector, location_vector)
n_vector = np.cross(origin_vector, location_vector)
s = np.linalg.norm(n_vector)
print(c,s)
n_vector_invert = np.array((
[0,-n_vector[2],n_vector[1]],
[n_vector[2],0,-n_vector[0]],
[-n_vector[1],n_vector[0],0]
))
I = np.eye(3)
# 核心公式:见上图
R_w2c = I + n_vector_invert + np.dot(n_vector_invert, n_vector_invert)/(1+c)
print(R_w2c)