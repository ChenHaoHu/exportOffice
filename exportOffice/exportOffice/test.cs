using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace exportexcel
{
    class test
    {

        static void Main(string[] args)
        {
            //test exportExcel
            // Excel exportExcel = new Excel();
            // exportExcel.exportExcel();

            //test exportWord
            //Word exportWord = new Word();
            // exportWord.exportWord();

            //test exportPPT
            PowerPoint exportPPT = new PowerPoint();
            exportPPT.exportPPT();
        }
       
    }
}
