using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace MinimalExcel
{
    class ExcelControl
    {
        ExcelControl()
        {
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = true;
            excelApp.Workbooks.Add();

            Excel._Worksheet mySheet = (Excel.Worksheet) excelApp.ActiveSheet;
            mySheet.Cells[1, "A"] = "Hallo Welt";
            mySheet.Cells[2, "A"] = "1";
            mySheet.Cells[3, "A"] = "=A2+2";

            var workSheet_range = mySheet.get_Range("A1", "B3");
            workSheet_range.Borders[Excel.XlBordersIndex.xlEdgeBottom].Color = ConsoleColor.Green;
			
			// test KOmmentar

            // https://docs.microsoft.com/de-de/dotnet/csharp/programming-guide/interop/walkthrough-office-programming

            // Console.WriteLine("Excel App erzeugt!");
            // Console.ReadKey();

        }


        static void Main(string[] args)
        {
            new ExcelControl();
        }
    }
}
