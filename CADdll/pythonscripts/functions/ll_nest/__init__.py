

import academit
from .minkowski import llnest_mksum, llnest_mkdirect, llnest_mkconnect, llnest_mkconvexhull, llnest_mkminibound

def 命令():  
    academit.添加命令("llnest-mksum", llnest_mksum)  
    academit.添加命令("llnest-mkdirect", llnest_mkdirect)  
    academit.添加命令("llnest-mkconnect", llnest_mkconnect)  
    academit.添加命令("llnest-mkconvexhull", llnest_mkconvexhull)  
    academit.添加命令("llnest-mkminibound", llnest_mkminibound)  