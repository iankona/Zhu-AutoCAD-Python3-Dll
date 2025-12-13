using System;

namespace MyFirstPlugin
{
    public class MyfirstAttribute:Attribute
    {
        public MyfirstAttribute(string obj1)
        {
            Console.WriteLine("特性1");
        }
        public MyfirstAttribute(string obj1, int obj2)
        {
            Console.WriteLine("特性2");
        }
        public MyfirstAttribute(string obj1, int obj2, float obj3)
        {
            Console.WriteLine("特性3");
        }

    }


    public delegate void 委托类型();
    public class HelloWorld:Attribute
    {
        public static 委托类型 action;
        public static void Set委托(ref Object obj) 
        {
            HelloWorld.action = (委托类型)obj;
        }

        public void 函数()
        {
            HelloWorld.action();
        }
    }
}
