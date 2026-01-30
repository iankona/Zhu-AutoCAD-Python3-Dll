import clr

import acad
import academit

import System

import os
import openpyxl as Excel
import subprocess



def 命令():  
    academit.添加命令("llarea-rec-for", llarea_rec_for)  
    academit.添加命令("llarea-rec-for-save-excel", llarea_rec_for_save_excel)  



zhu_text_size = 100


@acad.decorator_command
def llarea_rec_for():
    global zhu_text_size
    size = acad.GetInt(zhu_text_size, "请输入文字大小:")
    if size == None: size = 50
    zhu_text_size = size

    objidlist = acad.SSGetIdList([[0, "LWPOLYLINE"]]) 
    with acad.transaction() as trans:
        sumarea = 0.0 
        for i, objid in enumerate(objidlist):
            pt1 = acad.GetEntityBoundCenterXY0(objid)
            area = acad.GetEntityArea(objid) / 1000000
            sumarea += area
            char = f"{i+1}面积{area:.2f}平米"
            acad.AddText(pt1, char, zhu_text_size)
        pt1 = acad.GetPoint("请点击总面积放置位置:")
        char = f"共{i+1}个总面积{sumarea:.2f}平米"
        acad.AddText(pt1, char, zhu_text_size)




def cmd(command):
    subp = subprocess.Popen(command, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, encoding="utf-8")
    subp.wait(30) #  等待30秒再查询结果
    if subp.poll() == 0:
        print(subp.communicate()[1])
    else:
        print("失败")









@acad.decorator_command
def llarea_rec_for_save_excel():
    global zhu_text_size
    size = acad.GetInt(zhu_text_size, "请输入文字大小:")
    if size == None: size = 50
    zhu_text_size = size

    objidlist = acad.SSGetIdList([[-4, "<OR"],[0, "LWPOLYLINE"],[0, "*TEXT"], [0, "CIRCLE"], [-4, "OR>"]]) 
    with acad.transaction() as trans:
        textlist, plinelist = [], []
        for objid in objidlist:
            objref = acad.TransObjectForRead(objid)
            objtype = str(objref.GetType())
            if objtype == "Autodesk.AutoCAD.DatabaseServices.DBText":
                string = objref.TextString
                extend = objref.GeometricExtents
                point1 = extend.MinPoint
                point2 = extend.MaxPoint
                center = [(point1.X+point2.X)/2, (point1.Y+point2.Y)/2, 0]
                textlist.append([center, objid, string])
            if objtype == "Autodesk.AutoCAD.DatabaseServices.Polyline" or objtype == "Autodesk.AutoCAD.DatabaseServices.Circle": 
                if objref.Closed == False: continue
                area = objref.Area / 1000000
                extend = objref.GeometricExtents
                point1 = extend.MinPoint
                point2 = extend.MaxPoint
                width  = point2.Y - point1.Y
                length = point2.X - point1.X
                center = [(point1.X+point2.X)/2, (point1.Y+point2.Y)/2, 0]
                plinelist.append([center, objid, length, width, area])

    check_index_list = []
    for k, pvaluelist in enumerate(plinelist):
        objid = plinelist[k][1]
        for i, tvaluelist in enumerate(textlist):
            if i in check_index_list: continue
            pt1 = tvaluelist[0]
            flag = acad.IsPointInRange(pt1, objid)
            if flag:
                plinelist[k].append(tvaluelist[2])
                check_index_list.append(i)


    table = [[["序号", "名称","长", "宽", "计算面积","实际面积", "条件面积", "备注"], "项目名称"]]
    limitsumarea, sumarea = 0.0, 0.0
    for j, pvaluelist in enumerate(plinelist):
        center, objid, length, width, area = pvaluelist[:5]
        string = ""
        for char in pvaluelist[5:]:
            string += "\n"
            string += char
        if string != "": string = string[1:]
        limitarea = area
        if limitarea < 0.6: limitarea = 0.6
        sumarea += area
        limitsumarea += limitarea
        table.append([[j+1, 0, length, width, 0, area, limitarea], string])

    with acad.transaction() as trans:
        for m, pvaluelist in enumerate(plinelist):
            pt1, area = pvaluelist[0], pvaluelist[4]
            char = f"{m+1}面积{area:.2f}平米"
            acad.AddText(pt1, char, zhu_text_size)
        pt1 = acad.GetPoint()
        char = f"共{m+1}个总面积{sumarea:.2f}平米"
        acad.AddText(pt1, char, zhu_text_size)

    # 创建工作簿
    workbook = Excel.Workbook()
    sheet = workbook.active
    # sheet.append(table) # 写入多行，混合格式，1行字符串，1行数值会出错
    for row, string in table: sheet.append(row) # 写入行
    for n, [row, string] in enumerate(table): sheet[f"B{n+1}"] = string 

    filepath = acad.doc.Name
    # os.path.split(filepath)    # F:\U桌面备份\Desktop, Drawing1.dxf
    # os.path.splitext(filepath) # F:\U桌面备份\Desktop\Drawing1, .dxf
    pathhead, pattail = os.path.splitext(filepath)
    savepath = pathhead+'.xlsx'
    workbook.save(filename=savepath)
    # workbook.save(filename=pathhead+'.xlsx')
    # cmd("java -version")
    # cmd("exit 1")
    exepath = "\""+savepath+"\""
    xlspath = "\"F:\\U桌面备份\\Desktop\\Drawing1.xlsx\""
    cmd(exepath + " " +xlspath)


# DBText text1 = new DBText(); // 新建单行文本对象
# text1.Position = pt1[1];
# text1.TextString = "你好 数据智能笔记！text1";
# text1.Height = 10;
# text1.Rotation = Math.PI * 0.1;
# text1.IsMirroredInY = true; // 在Y轴镜像
# text1.HorizontalMode = TextHorizontalMode.TextLeft;
# text1.AlignmentPoint = text1.Position; // 设置对齐点
# db.AddEntityToModeSpace(text1);