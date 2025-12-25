<div align="center" style="font-size: 42px; font-weight:bold;">AutoCAD2023Python3.13ModDll</div>

### 

### 感恩老天爷

感恩老天爷恩赐，可以在AutoCAD里用命令调用AutoCADLisp的方式，用命令调用python代码了。

解决方案思路AutoCAD的netload命令+pythonnet3.05+python3.13。

实现分三步，第1步，用netload加载pythonnet的dll和python的dll，第2步，用python调用net的Emit生成插件dll，第3步，再用netload加载生成的插件dll把命令注册到CAD软件里面，并用net的Assembly打开生成的插件dll，在把第1步的命令对应的python函数赋值给插件dll，这样autocad就能正常用命令调用python函数了，也能使用Net的lockdoc+db绘制图纸的方式了。

已知问题，只能打开1个cad，再打开第2个cad程序，第二个程序崩溃，而第一个cad主体程序正常。





