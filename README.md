<div align="center" style="font-size: 42px; font-weight:bold;">AutoCAD2023Python3.13ModDll</div>

### 

### 感恩老天爷

感恩老天爷恩赐，可以在AutoCAD里用命令调用AutoCADLisp的方式，用命令调用python代码了。

解决方案思路AutoCAD的netload命令+pythonnet3.05+python3.13。

实现分三步，第1步，用netload加载pythonnet的dll和python的dll，第2步，用python调用net的Emit生成插件dll，第3步，再用netload加载生成的插件dll把命令注册到CAD软件里面，并用net的Assembly打开生成的插件dll，在把第1步的命令对应的python函数赋值给插件dll，这样autocad就能正常用命令调用python函数了。

已知问题，只能打开1个cad，再打开第2个cad程序，第二个程序崩溃，而第一个cad主体程序正常。

可以是楼主对pythonnet不熟悉，楼主觉得pythonnet 好像没有在python里支持net的using函数域语法，和 is  as （TableBlock）等net的强制类型转换，这个缺失导致很多cad的net函数无法使用，使用这些net函数常常需要using 和 as 一起使用，楼主尝试着直接调用使用这些函数，结果cad直接崩溃。除去这些，也有很多的cad的net的C#类型和函数能用的，也能实现很大的一部分，用python代码直接操作cad了。

另1个用python操作cad的方式是pythoncom，但楼主好像没有找到，pythoncom把python函数注册进cad里，成为在cad里面可以用命令就可以直接调用python代码的方式。





