using MSWord = Microsoft.Office.Interop.Word;

using System.IO;

using System.Reflection;
using System;


    class Word
    {
        public void exportWord()

        {

            object path;                          //文件路径变量

            string strContent;                    //文本内容变量

            MSWord.Application wordApp;               //Word应用程序变量

            MSWord.Document wordDoc;              //Word文档变量



            path = @"D:\MyWord.doc";              //路径

            wordApp = new MSWord.ApplicationClass(); //初始化

            //如果已存在，则删除

            if (File.Exists((string)path))

            {

                File.Delete((string)path);

            }

            //由于使用的是COM库，因此有许多变量需要用Missing.Value代替

            Object Nothing = Missing.Value;

            wordDoc = wordApp.Documents.Add(ref Nothing, ref Nothing, ref Nothing, ref Nothing);

        //写入普通文本

        strContent = "普通文本普通文本普通文本普通文本普通文本\n";

        wordDoc.Paragraphs.Last.Range.Text = strContent;

        //写入黑体文本

        strContent = "黑体文本黑体文本黑体文本黑体文本黑体文本\n";

        wordDoc.Paragraphs.Last.Range.Font.Name = "黑体";

        wordDoc.Paragraphs.Last.Range.Text = strContent;

        //写入加粗文本

        strContent = "加粗文本加粗文本加粗文本加粗文本加粗文本\n";

        wordDoc.Paragraphs.Last.Range.Font.Bold = 1;

        wordDoc.Paragraphs.Last.Range.Text = strContent;

        //写入15号字体文本

        strContent = "15号字体文本15号字体文本15号字体文本15号字体文本\n";

        wordDoc.Paragraphs.Last.Range.Font.Size = 15;

        wordDoc.Paragraphs.Last.Range.Text = strContent;

        //写入斜体文本

        strContent = "斜体文本斜体文本斜体文本斜体文本斜体文本\n";

        wordDoc.Paragraphs.Last.Range.Font.Italic = 1;

        wordDoc.Paragraphs.Last.Range.Text = strContent;

        //写入蓝色文本

        strContent = "蓝色文本蓝色文本蓝色文本蓝色文本蓝色文本\n";

        wordDoc.Paragraphs.Last.Range.Font.Color = MSWord.WdColor.wdColorBlue;

        wordDoc.Paragraphs.Last.Range.Text = strContent;

        //写入下画线文本

        strContent = "下画线文本下画线文本下画线文本下画线文本下画线文本\n";

        wordDoc.Paragraphs.Last.Range.Font.Underline = MSWord.WdUnderline.wdUnderlineThick;

        wordDoc.Paragraphs.Last.Range.Text = strContent;

        //写入红色下画线文本

        strContent = "红色下画线文本红色下画线文本红色下画线文本红色下画线文本\n";

        wordDoc.Paragraphs.Last.Range.Font.Underline = MSWord.WdUnderline.wdUnderlineThick;

        wordDoc.Paragraphs.Last.Range.Font.UnderlineColor = MSWord.WdColor.wdColorRed;

        wordDoc.Paragraphs.Last.Range.Text = strContent;

        //定义一个Word中的表格对象

        MSWord.Table table = wordDoc.Tables.Add(wordApp.Selection.Range, 5, 5, ref Nothing, ref Nothing);

        //默认创建的表格没有边框，这里修改其属性，使得创建的表格带有边框

        table.Borders.Enable = 1;

        //使用两层循环填充表格的内容

        for (int i = 1; i <= 5; i++)

        {

            for (int j = 1; j <= 5; j++)

            {

                table.Cell(i, j).Range.Text = "第" + i + "行，第" + j + "列";

            }

        }

        //图片文件的路径

        string filename = @"D:\1.png";

        //要向Word文档中插入图片的位置

        Object range = wordDoc.Paragraphs.Last.Range;

        //定义该插入的图片是否为外部链接

        Object linkToFile = false;                //默认

        //定义要插入的图片是否随Word文档一起保存

        Object saveWithDocument = true;               //默认

        //使用InlineShapes.AddPicture方法插入图片

        wordDoc.InlineShapes.AddPicture(filename, ref linkToFile, ref saveWithDocument, ref range);



        //WdSaveFormat为Word文档的保存格式

        object format = MSWord.WdSaveFormat.wdFormatDocument;

            //将wordDoc文档对象的内容保存为DOC文档

            wordDoc.SaveAs(ref path, ref format, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);

            //关闭wordDoc文档对象

            wordDoc.Close(ref Nothing, ref Nothing, ref Nothing);

            //关闭wordApp组件对象

            wordApp.Quit(ref Nothing, ref Nothing, ref Nothing);

            Console.WriteLine(path + " 创建完毕！");

        }

    }

