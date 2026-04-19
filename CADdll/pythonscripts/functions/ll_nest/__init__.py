

import academit
from .minkowski import llnest_mksum, llnest_mkdirect, llnest_mkconnect, llnest_mkconvexhull, llnest_mkminibound

def 命令():  
    academit.添加命令("llnest-mksum", llnest_mksum)  
    academit.添加命令("llnest-mkdirect", llnest_mkdirect)  
    academit.添加命令("llnest-mkconnect", llnest_mkconnect)  
    academit.添加命令("llnest-mkconvexhull", llnest_mkconvexhull)  
    academit.添加命令("llnest-mkminibound", llnest_mkminibound)  




# 矩形装箱算法
# def pack(rects):
# rects = sorted(rects, key=lambda r: max(r))
# bins = []
# while rects:
# bin = []
# width_left = BOUND
# height = 0
# for rect in rects:
# if rect[0]<= width_left:
# bin.append(rect)
# width_left -= rect[0]
# height = max(height, rect[1])
# for rect in bin:
# rects.remove(rect)
# bins.append((BOUND - width_left, height, bin))
# return bins
# 以上是一个简单的矩形排料算法实现。代码中首先将矩形按照最大边长从大到小排序。然后依次放入宽度为BOUND的“箱子”中。在每个箱子中，从剩余宽度中选取最大高度的矩形，直到无法再放入为止。
# 这段代码的复杂度为O(N^2)，对于大规模问题可能会比较慢。但对于小规模问题，这种简单的贪心算法已经能够得到较好的结果。