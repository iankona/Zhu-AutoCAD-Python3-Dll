

# 以下是一个射线法的Python实现：
def is_point_in_polygon(pt1, mkptlist):
    x, y = pt1[0:2]
    inside = False
    p1x, p1y = mkptlist[0][0:2]
    for pt2 in mkptlist[:]+mkptlist[0:1]:
        p2x, p2y = pt2[0:2]
        if y > min(p1y, p2y):
            if y <= max(p1y, p2y):
                if x <= max(p1x, p2x):
                    if abs(p1y-p2y) > 0.00001: # p1y != p2y
                        xinters = (y - p1y) * (p2x - p1x) / (p2y - p1y) + p1x
                    if abs(p1x-p2x) < 0.00001 or x <= xinters: # p1y == p2y
                        inside = not inside
        p1x, p1y = p2x, p2y
    return inside






























# 以下是一个射线法的Python实现：
# def is_point_in_polygon(point, polygon):

#     x, y = point


#     n = len(polygon)


#     inside = False


#     p1x, p1y = polygon[0]


#     for i in range(n + 1):


#         p2x, p2y = polygon[i % n]


#         if y > min(p1y, p2y):


#             if y <= max(p1y, p2y):


#                 if x <= max(p1x, p2x):


#                     if p1y != p2y:


#                         xinters = (y - p1y) * (p2x - p1x) / (p2y - p1y) + p1x


#                     if p1x == p2x or x <= xinters:


#                         inside = not inside


#         p1x, p1y = p2x, p2y


#     return inside


# 定义多边形的顶点

# polygon_points = [(0, 0), (1, 0), (1, 1), (0, 1)]


# 定义需要判断的点

# point = (0.5, 0.5)


# 判断点是否在多边形内

# is_inside = is_point_in_polygon(point, polygon_points)


# print(is_inside)  # 输出：True