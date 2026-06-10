
import acad
import academit

# import open3d as o3d
# import numpy as np

def 命令(): 
    academit.添加命令("ll-point-cloud-load", ll_point_cloud_load)
    pass


@acad.decorator_command
def ll_point_cloud_load():
    pcd = o3d.io.read_point_cloud(r"f:\NetPythonProject\PyKinect2-PyQtGraph-PointClouds\models\zhu_cloud_701_002.pcd")
    points = np.asarray(pcd.points)
    colors = np.asarray(pcd.colors)*255



    with acad.transaction() as trans:
        # count = 0
        # countsum = len(points)
        for pt1, [r,g,b] in zip(points, colors):
            acad.AddPointCloud(pt1, [int(r), int(g), int(b)])
            # count += 1
            # if count % 10000 == 0: 
            #     acad.Prompt(f"{count/countsum}\n")

    # o3d.visualization.draw_geometries([pcd])


