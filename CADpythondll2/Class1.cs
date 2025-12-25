using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using Python.Runtime;
using System;
using System.Collections.Generic;


[assembly: CommandClass(typeof(ZhuACADPlugin.HelloWorld))]
namespace ZhuACADPlugin
{
    public class HelloWorld
    {
        public static bool flag1 = false;
        public static bool flag2 = false;

        [CommandMethod("Zhu_Hello")]
        public void 函数1()
        {
            if (flag1 == true) { return; }

            Document doc = Application.DocumentManager.MdiActiveDocument;
            doc.Editor.WriteMessage("...Hello ");

            Runtime.PythonDLL = "E:\\CADdll\\python-3.13.11-embed-amd64\\python313.dll";
            PythonEngine.Initialize();
            using (Py.GIL())
            {
                dynamic sys = Py.Import("sys");
                sys.path.append("E:\\CADdll\\pythonscripts");
                sys.path.append("E:\\CADdll\\pythonscripts\\libs");
                sys.path.append("E:\\CADdll\\pythonscripts\\functions");
                dynamic manage = Py.Import("manage");
                manage.生成命令();

            }

            flag1 = true;

            doc.Editor.WriteMessage("Python...");
        }

        [CommandMethod("Zhu_World")]
        public void 函数2()
        {
            if (flag1 == false) { return; }
            if (flag2 == true ) { return; }

            Document doc = Application.DocumentManager.MdiActiveDocument;

            using (Py.GIL())
            {
                dynamic manage = Py.Import("manage");
                manage.设置命令();

            }

            flag2 = true;

            doc.Editor.WriteMessage("...Hello World...");
        }

    }
}