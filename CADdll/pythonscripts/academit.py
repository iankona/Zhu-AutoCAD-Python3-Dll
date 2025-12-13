import clr

import System
from System.Reflection import Assembly, AssemblyName, TypeAttributes, FieldAttributes, MethodAttributes, CallingConventions, PropertyAttributes, ParameterAttributes
from System.Reflection.Emit import AssemblyBuilderAccess, AssemblyBuilder, ModuleBuilder, TypeBuilder, FieldBuilder, MethodBuilder, PropertyBuilder, ConstructorBuilder, ILGenerator, OpCodes, ParameterBuilder, LocalBuilder, CustomAttributeBuilder

# from Autodesk.AutoCAD.Runtime import CommandClassAttribute, CommandMethodAttribute # accoremgd.dll

assemly = Assembly.LoadFile("E:\\AutoCAD\AutoCAD 2023\\accoremgd.dll")
classattribute = assemly.GetType("Autodesk.AutoCAD.Runtime.CommandClassAttribute")
methodattribute = assemly.GetType("Autodesk.AutoCAD.Runtime.CommandMethodAttribute")

# print(System.Type.GetType("Autodesk.AutoCAD.Runtime.CommandClassAttribute")) # None # 不能用
# print(System.Type.GetType("Autodesk.AutoCAD.Runtime.CommandMethodAttribute")) # None # 不能用

程序集名称 = "ZhuAcadPlugin"
程序集 = System.AppDomain.CurrentDomain.DefineDynamicAssembly(AssemblyName(程序集名称), AssemblyBuilderAccess.RunAndSave)
模块 = 程序集.DefineDynamicModule(程序集名称, 程序集名称 + ".dll")

def 保存程序集():
    程序集.Save(程序集名称 + ".dll")
    # 程序集.Save(f"F:\\CADdll\\{程序集名称}.dll")
    print(f"保存程序集：{程序集名称}.dll")


def 测试2():
    print("鹅鹅鹅222222222!!!!")


类型名称列表 = []
类型字典 = {}


def 添加命令(命令:str, Python函数):
    类型名称 = "Zhu_" + 命令.replace("-", "_")
    字段名称 = "action"
    委托类型 = System.Action
    委托实例 = System.Action(Python函数)
    设置名称 = "SetAction"
    方法名称 = "函数"

    if 类型名称 in 类型名称列表: raise ValueError("CAD命令:{命令}重复......")
    类型名称列表.append(类型名称)

    类型 = 模块.DefineType(类型名称, TypeAttributes.Public)
    特性构造 = classattribute.GetConstructor([System.Type])
    特性 = CustomAttributeBuilder(特性构造, [类型])
    类型.SetCustomAttribute(特性)

    字段 = 类型.DefineField(字段名称, 委托类型, FieldAttributes.Public| FieldAttributes.Static) 

    设置 = 类型.DefineMethod(设置名称, MethodAttributes.Public|MethodAttributes.Static, System.Void, [System.Object().GetType().MakeByRefType()]) 
    IL = 设置.GetILGenerator()
    IL.Emit(OpCodes.Ldarg_0)
    IL.Emit(OpCodes.Ldind_Ref)   
    IL.Emit(OpCodes.Castclass, 委托类型)
    IL.Emit(OpCodes.Stsfld, 字段)
    IL.Emit(OpCodes.Ret)

    方法 = 类型.DefineMethod(方法名称, MethodAttributes.Public, System.Void, []) 
    特性构造 = methodattribute.GetConstructor([System.String])
    特性 = CustomAttributeBuilder(特性构造, [命令])
    方法.SetCustomAttribute(特性)
    IL = 方法.GetILGenerator()
    label = IL.DefineLabel()
    IL.Emit(OpCodes.Ldsfld, 字段)
    IL.Emit(OpCodes.Brfalse_S, label)
    IL.Emit(OpCodes.Ldsfld, 字段)
    IL.Emit(OpCodes.Callvirt, 委托实例.GetType().GetMethod("Invoke"))
    IL.MarkLabel(label)
    IL.Emit(OpCodes.Ret)

    classtype = 类型.CreateType()
    # classtype.GetMethod("SetAction").Invoke(System.Void,[委托实例]) 
    # instance = System.Activator.CreateInstance(classtype)
    # instance.函数() 
    类型字典[类型名称] = [classtype, 委托实例]


# 添加命令("LL-HelloWorld2", 测试2)
# 保存程序集()


def 设置程序集():
    assemly = Assembly.LoadFile(f"F:\\CADdll\\{程序集名称}.dll")
    for typename, [classtype, 委托实例] in 类型字典.items():
        classtype = assemly.GetType(typename)
        classtype.GetMethod("SetAction").Invoke(System.Void,[委托实例]) 

