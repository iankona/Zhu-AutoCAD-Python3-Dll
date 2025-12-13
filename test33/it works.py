import clr
clr.AddReference("clrclass")


import System
from MyFirstPlugin import HelloWorld, 委托类型
from System.Reflection import Assembly, AssemblyName, TypeAttributes, FieldAttributes, MethodAttributes, CallingConventions, PropertyAttributes, ParameterAttributes
from System.Reflection.Emit import AssemblyBuilderAccess, AssemblyBuilder, ModuleBuilder, TypeBuilder, FieldBuilder, MethodBuilder, PropertyBuilder, ConstructorBuilder, ILGenerator, OpCodes, ParameterBuilder, LocalBuilder

assemly = Assembly.LoadFile("C:\\Users\\Administrator\\test22\\clrclass.dll")

methods = assemly.GetType("MyFirstPlugin.委托类型").GetMethods()



invoke = assemly.GetType("MyFirstPlugin.委托类型").GetMethod("Invoke")


# for method in methods:
#     print(method)


# methods = assemly.GetType("MyFirstPlugin.HelloWorld").GetMethods()
# for method in methods:
#     print(method)


程序集名称 = "Zhu测试函数"
程序集 = System.AppDomain.CurrentDomain.DefineDynamicAssembly(AssemblyName(程序集名称), AssemblyBuilderAccess.RunAndSave)
模块 = 程序集.DefineDynamicModule(程序集名称, 程序集名称 + ".dll")

def 保存程序集():
    程序集.Save(程序集名称 + ".dll")
    print(f"保存程序集：{程序集名称}.dll")



def 测试2():
    print("鹅鹅鹅222222222！！！！")



def 添加命令():

    类型 = 模块.DefineType("HelloWorld", TypeAttributes.Public)

    字段 = 类型.DefineField("func", 委托类型, FieldAttributes.Public| FieldAttributes.Static) 

    设置 = 类型.DefineMethod("Set委托", MethodAttributes.Public|MethodAttributes.Static, System.Void, [System.Object().GetType().MakeByRefType()]) 
    IL = 设置.GetILGenerator()
    IL.Emit(OpCodes.Ldarg_0)
    IL.Emit(OpCodes.Ldind_Ref)   
    IL.Emit(OpCodes.Castclass, 委托类型)
    IL.Emit(OpCodes.Stsfld, 字段)
    IL.Emit(OpCodes.Ret)

    方法 = 类型.DefineMethod("函数", MethodAttributes.Public, System.Void, []) 
    IL = 方法.GetILGenerator()
    IL.Emit(OpCodes.Ldsfld, 字段)
    IL.Emit(OpCodes.Callvirt, invoke)
    IL.Emit(OpCodes.Ret)

    classtype = 类型.CreateType()
    classtype.GetMethod("Set委托").Invoke(System.Void,[委托类型(测试2)]) 
    instance = System.Activator.CreateInstance(classtype)
    instance.函数() 


添加命令()
保存程序集()