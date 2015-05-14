using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;


namespace Read_From_Excel_1_
{
    class Program
    {
        static void Main(string[] args)
        {
            ApplicationClass appExcel = new ApplicationClass();
            Workbook newWorkBook = appExcel.Workbooks.Open("C:\\Users\\Udokoro\\Desktop\\APPS NAME, LOGIN DETAILS AND RESOURCES.xls", true, true);
            _Worksheet objSheet = (_Worksheet)appExcel.ActiveWorkbook.ActiveSheet;
            string value = "James";
            for (int i = 1; i < 5; i++)
            {
                value = objSheet.get_Range("A"+i).get_Value().ToString();
                System.Console.WriteLine("This is the value in the cell A1 - " + value);
            }

            Console.WriteLine("");

            Range range = objSheet.get_Range("$A$1:$A$4");
            object[,] input = (object[,])range.Cells.Value;
            string[] testing = input.Cast<string>().ToArray();             
                    

            foreach(string s in testing)
            {
                Console.WriteLine(s);
            }

            System.Console.ReadLine();

        }
    }
}
