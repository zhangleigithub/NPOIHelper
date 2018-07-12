using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOIEx;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace NpoiHelper
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream file = new FileStream(@"C:\Users\zhanglei\Desktop\01.xls", FileMode.Open, FileAccess.Read))
            {
                HSSFWorkbook hssfWorkBook = new HSSFWorkbook(file);

                //sheet
                ISheet sheet1 = hssfWorkBook.GetSheet("Sheet1");
                var var = sheet1.ConvertSheetToObjects<Person>();
            }
        }
    }
}
