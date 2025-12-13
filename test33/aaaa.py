import clr
clr.AddReference("clrclass")

import System
from System.Reflection import Assembly, AssemblyName, TypeAttributes, FieldAttributes, MethodAttributes, CallingConventions, PropertyAttributes, ParameterAttributes
from System.Reflection.Emit import AssemblyBuilderAccess, AssemblyBuilder, ModuleBuilder, TypeBuilder, FieldBuilder, MethodBuilder, PropertyBuilder, ConstructorBuilder, ILGenerator, OpCodes, ParameterBuilder, LocalBuilder, CustomAttributeBuilder

# from MyFirstPlugin import MyfirstAttribute

# 方式2 能用
# assemly = Assembly.LoadFile("C:\\Users\\Administrator\\test22\\clrclass.dll")
# methods = assemly.GetType("MyFirstPlugin.委托类型").GetMethods()
# invoke = assemly.GetType("MyFirstPlugin.委托类型").GetMethod("Invoke")


程序集名称 = "Zhu测试函数"
程序集 = System.AppDomain.CurrentDomain.DefineDynamicAssembly(AssemblyName(程序集名称), AssemblyBuilderAccess.RunAndSave)
模块 = 程序集.DefineDynamicModule(程序集名称, 程序集名称 + ".dll")

def 保存程序集():
    程序集.Save(程序集名称 + ".dll")
    print(f"保存程序集：{程序集名称}.dll")


def 测试2():
    print("鹅鹅鹅222222222!!!!")


类型名称列表 = []
类型列表 = []


def 添加命令(命令:str, Python函数):
    类型名称 = "Zhu_" + 命令.replace("-", "_")
    字段名称 = "action"
    委托类型 = System.Action
    委托实例 = System.Action(Python函数)
    设置名称 = "SetAction"
    # 特性名称 = "MyFirstPlugin.MyfirstAttribute"
    # 特性类型 = assemly.GetType(特性名称)
    方法名称 = "函数"

    if 类型名称 in 类型名称列表: raise ValueError("CAD命令:{命令}重复......")
    类型名称列表.append(类型名称)

    类型 = 模块.DefineType(类型名称, TypeAttributes.Public)
    # 特性构造 = MyfirstAttribute.GetType().GetConstructor([System.String])
    # 特性 = CustomAttributeBuilder(特性构造, [命令])
    # 类型.SetCustomAttribute(特性)

    字段 = 类型.DefineField(字段名称, 委托类型, FieldAttributes.Public| FieldAttributes.Static) 

    设置 = 类型.DefineMethod(设置名称, MethodAttributes.Public|MethodAttributes.Static, System.Void, [System.Object().GetType().MakeByRefType()]) 
    IL = 设置.GetILGenerator()
    IL.Emit(OpCodes.Ldarg_0)
    IL.Emit(OpCodes.Ldind_Ref)   
    IL.Emit(OpCodes.Castclass, 委托类型)
    IL.Emit(OpCodes.Stsfld, 字段)
    IL.Emit(OpCodes.Ret)

    方法 = 类型.DefineMethod(方法名称, MethodAttributes.Public, System.Void, []) 
    # 特性构造 = 特性类型.GetConstructor([System.String])
    # 特性 = CustomAttributeBuilder(特性构造, [命令])
    # 方法.SetCustomAttribute(特性)
    IL = 方法.GetILGenerator()
    label = IL.DefineLabel()
    IL.Emit(OpCodes.Ldsfld, 字段)
    IL.Emit(OpCodes.Brfalse_S, label)
    IL.Emit(OpCodes.Ldsfld, 字段)
    IL.Emit(OpCodes.Callvirt, 委托实例.GetType().GetMethod("Invoke"))
    IL.MarkLabel(label)
    IL.Emit(OpCodes.Ret)

    classtype = 类型.CreateType()
    classtype.GetMethod(设置名称).Invoke(System.Void,[委托实例]) 
    instance = System.Activator.CreateInstance(classtype)
    instance.函数() 
    类型列表.append(classtype)

添加命令("LL-HelloWorlld", 测试2)
保存程序集()