using MSExcel = Microsoft.Office.Interop.Excel;

using System.IO;
using exportexcel;
using System.Reflection;
using System;
using Microsoft.Office.Core;
using System.Collections.Generic;


class Excel

{
    private static int rowIndex = 18;
    //此处是数据条数加7
    private static int colIndex = 19;

    public  void exportExcel()

    {

        object path;                          //文件路径变量

        MSExcel.Application excel;              //Excel应用程序变量

        MSExcel.Workbook excelDoc;                     //Excel文档变量



        path = @"D:\MyExcel.xlsx";                    //路径

        excel = new MSExcel.ApplicationClass();    //初始化

        //如果已存在，则删除

        if (File.Exists((string)path))

        {

            File.Delete((string)path);

        }

        //由于使用的是COM库，因此有许多变量需要用Nothing代替

        Object Nothing = Missing.Value;

        excelDoc = excel.Workbooks.Add(Nothing);

        //使用第一个工作表作为插入数据的工作表

        MSExcel.Worksheet xSt = (MSExcel.Worksheet)excelDoc.Sheets[1];



       
        combineTransverse(1,1, rowIndex, "辽宁石油化工大学有限公司", excel,xSt);
        combineTransverse(2, 1, 10, "分析项目：过程示例", excel, xSt);
        combineTransverse(2, 11, rowIndex, "表页：1/2", excel, xSt);
        combineTransverse(3, 1, 2, "图纸编号", excel, xSt);
        combineTransverse(4, 1, 2, "小组成员", excel, xSt);
        combineTransverse(5, 1, 2, "节点", excel, xSt);
        combineTransverse(6, 1, 2, "设计意图", excel, xSt);
        combineTransverse(3, 3, 15, "图纸编号明细", excel, xSt);
        combineTransverse(4, 3, 15, "小组成员明细", excel, xSt);
        combineTransverse(5, 3, 15, "节点明细", excel, xSt);
        combineTransverse(3, 16, rowIndex, "日期：2018/5/2", excel, xSt);
        combineTransverse(4, 16, rowIndex, "会议日期：2018/5/2", excel, xSt);
        combineTransverse(5, 16, rowIndex, "  ", excel, xSt);
      
        combineTransverseVertical(6, 1, 7, 2, "设计意图明细", excel, xSt);
        combineTransverse(6, 3, 4, "起料", excel, xSt);
        combineTransverse(7, 3, 4, "起始点", excel, xSt);
        combineTransverse(6, 5, 10, "起料详细", excel, xSt);
        combineTransverse(7, 5, 10, "起始点详细", excel, xSt);
        combineTransverse(6, 11, 12, "活动", excel, xSt);
        combineTransverse(7, 11, 12, "终止点", excel, xSt);
        combineTransverse(6, 13, 18, "活动详细", excel, xSt);
        combineTransverse(7, 13, 18, "终止点详细", excel, xSt);


        //data
        combineTransverse(8, 1, 1, "序号", excel, xSt);
        combineTransverse(8, 2, 2, "引导词", excel, xSt);
        combineTransverse(8, 3, 4, "要素", excel, xSt);
        combineTransverse(8, 5, 6, "偏离", excel, xSt);
        combineTransverse(8, 7, 8, "可能的原因", excel, xSt);
        combineTransverse(8, 9, 10, "后果", excel, xSt);
        combineTransverse(8, 11, 12, "安全措施", excel, xSt);
        combineTransverse(8, 13, 14, "注释", excel, xSt);
        combineTransverse(8, 15, 17, "建议措施", excel, xSt);
        combineTransverse(8, 18, 18, "责任人", excel, xSt);

        //生成测试数据
        List<data> datas = new List<data>();

        for(int i = 0; i < 10; i++)
        {
            datas.Add(new data(i, "guideword" + i,"key"+i,"deviate"+i, "possiblecause" + i, "consequence" + i, "safetymeasures" + i, "annotation" + i, "suggestionmeasure" + i, "responsibilityperson" + i));
        }

        //show data in excel
        foreach (data d in datas)
        {
            combineTransverse(9+d.Id, 1, 1, d.Id+" ", excel, xSt);
            combineTransverse(9 + d.Id, 2, 2, d.Guideword, excel, xSt);
            combineTransverse(9 + d.Id, 3, 4, d.Key, excel, xSt);
            combineTransverse(9 + d.Id, 5, 6, d.Deviate, excel, xSt);
            combineTransverse(9 + d.Id, 7, 8, d.Possiblecause, excel, xSt);
            combineTransverse(9 + d.Id, 9, 10, d.Consequence, excel, xSt);
            combineTransverse(9 + d.Id, 11, 12, d.Safetymeasures, excel, xSt);
            combineTransverse(9 + d.Id, 13, 14, d.Annotation, excel, xSt);
            combineTransverse(9 + d.Id, 15, 17, d.Suggestionmeasure, excel, xSt);
            combineTransverse(9 + d.Id, 18, 18, d.Responsibilityperson, excel, xSt);
        }

        // 
        //设置整个报表的标题格式 
        // 
        xSt.get_Range(excel.Cells[1, 1], excel.Cells[1, 1]).Font.Bold = true;
        xSt.get_Range(excel.Cells[1, 1], excel.Cells[1, 1]).Font.Size = 22;
    

        //设置报表表格为最适应宽度 
        // 
        xSt.get_Range(excel.Cells[1, 1], excel.Cells[rowIndex, colIndex]).Select();
        xSt.get_Range(excel.Cells[1, 1], excel.Cells[rowIndex, colIndex]).Columns.AutoFit();

        // 
        //绘制边框 
        // 
        xSt.get_Range(excel.Cells[1, 1], excel.Cells[rowIndex, colIndex-1]).Borders.LineStyle = 1;
        xSt.get_Range(excel.Cells[1, 1], excel.Cells[rowIndex, colIndex-1]).Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick;//设置左边线加粗 
        xSt.get_Range(excel.Cells[1, 1], excel.Cells[rowIndex, colIndex - 1]).Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick;//设置上边线加粗 
        xSt.get_Range(excel.Cells[1, 1], excel.Cells[rowIndex, colIndex - 1]).Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick;//设置右边线加粗 
        xSt.get_Range(excel.Cells[1, 1], excel.Cells[rowIndex, colIndex - 1]).Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick;//设置下边线加粗

        excel.Visible = true;


        //WdSaveFormat为Excel文档的保存格式

        object format = MSExcel.XlFileFormat.xlWorkbookDefault;

        //将excelDoc文档对象的内容保存为XLSX文档

        excelDoc.SaveAs(path, format, Nothing, Nothing, Nothing, Nothing, MSExcel.XlSaveAsAccessMode.xlExclusive, Nothing, Nothing, Nothing, Nothing, Nothing);

        //关闭excelDoc文档对象

        excelDoc.Close(Nothing, Nothing, Nothing);

        //关闭excelApp组件对象

        excel.Quit();

        Console.WriteLine(path + " 创建完毕！");

      //  Console.ReadLine();

    }


    //横向联合
    public  void combineTransverse(int row,int begin,int end,string text, MSExcel.Application excel, MSExcel.Worksheet xSt)
    {
        excel.Cells[row, begin] = text;
        //设置整个报表的标题为跨列居中 
        // 
        xSt.get_Range(excel.Cells[row, begin], excel.Cells[row, end]).Select();
        xSt.get_Range(excel.Cells[row, begin], excel.Cells[row, end]).HorizontalAlignment = XlHAlign.xlHAlignCenterAcrossSelection;
        xSt.get_Range(excel.Cells[row, begin], excel.Cells[row, end]).MergeCells = true;
        xSt.get_Range(excel.Cells[row, begin], excel.Cells[row, end]).WrapText = true;//


    }

    //多向联合
    public  void combineTransverseVertical(int begin1, int end1,  int begin2, int end2, string text, MSExcel.Application excel, MSExcel.Worksheet xSt)
    {
        excel.Cells[begin1, end1] = text;
        //设置整个报表的标题为跨列居中 
        xSt.get_Range(excel.Cells[begin1, end1], excel.Cells[begin2, end1]).MergeCells = true;
        xSt.get_Range(excel.Cells[begin1, end1], excel.Cells[begin2, end2]).MergeCells = true;
        xSt.get_Range(excel.Cells[begin1, end1], excel.Cells[begin2, end2]).WrapText = true;//  

    }
}