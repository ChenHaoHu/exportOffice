using System;

using System.Collections.Generic;

using System.Text;

using System.IO;

using PPT = Microsoft.Office.Interop.PowerPoint;

using System.Reflection;


class PowerPoint
{
    public void exportPPT()
    { 
            string path;                       //文件路径变量

            PPT.Application pptApp;                 //Excel应用程序变量

            PPT.Presentation pptDoc;                //Excel文档变量

            path = @"D:\MyPPT.ppt";                //路径

            pptApp = new PPT.ApplicationClass(); //初始化

            //如果已存在，则删除

            if (File.Exists((string)path))

            {

                File.Delete((string)path);

            }

            //由于使用的是COM库，因此有许多变量需要用Nothing代替

            Object Nothing = Missing.Value;

        pptDoc = pptApp.Presentations.Add(Microsoft.Office.Core.MsoTriState.msoFalse);

        pptDoc.Slides.Add(1, Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutBlank);

        string pic = @"D:\1.png";

        foreach (PPT.Slide slide in pptDoc.Slides)

        {

            slide.Shapes.AddPicture(pic, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, 150, 150, 300, 200);

        }

       

        //WdSaveFormat为PPTs文档的保存格式

        PPT.PpSaveAsFileType format = PPT.PpSaveAsFileType.ppSaveAsDefault;

        //将 pptDoc文档对象的内容保存为XLSX文档

        pptDoc.SaveAs(path, format, Microsoft.Office.Core.MsoTriState.msoFalse);

        //关闭pptDoc文档对象

        pptDoc.Close();

        //关闭pptApp组件对象

        pptApp.Quit();

            Console.WriteLine(path + " 创建完毕！");

           // Console.ReadLine();

        
    }
}

