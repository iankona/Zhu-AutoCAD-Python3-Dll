using System;

namespace MyFirstPlugin
{

    public class HelloWorld222
    {
        public static Action action = null;
        public void 函数()
        {
            if (HelloWorld.action != null)
            { HelloWorld.action(); }
        }
    }
}
