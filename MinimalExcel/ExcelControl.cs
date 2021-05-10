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
            // Objekt erzeugen
            Console.WriteLine("Erzeuge Excel COM Objekt");
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = true;

            // Öffnen einer Excel Datei
            Console.WriteLine("Öffne Datei");
            String filename = "ExcelFile.xlsx";
            String path = System.IO.Path.GetFullPath(filename);
            if (System.IO.File.Exists(path))
            {
                excelApp.Workbooks.Open(path);
            }
            Excel._Worksheet mySheet = (Excel.Worksheet)excelApp.ActiveSheet;

            // Lesen der Datei
            Console.WriteLine("Lese aus der Datei");
            Excel.Range bereich = mySheet.Cells[1, "A"] as Excel.Range;
            String zellWert = (String) bereich.Value;
            Console.WriteLine("   Wert: " + zellWert);

            Excel.Range workSheet_range = mySheet.get_Range("A1", "B3"); // Mehrere Zellen erhelten Sie auch so

            // Schreiben in die Datei
            Console.WriteLine("Schreibe in die Datei");
            mySheet.Cells[1, "B"] = "Hallo Welt";
            mySheet.Cells[2, "B"] = "1";
            mySheet.Cells[3, "B"] = "=A2+2";

            // Speichern der Datei und beenden.
            Console.WriteLine("Speichern der Änderungen");
            String newFileName = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(path), "newExcelFile.xlsx");
            mySheet.SaveAs(newFileName);
            excelApp.Quit();

            // Console.ReadKey();

        }


        static void Main(string[] args)
        {
            new ExcelControl();
        }
    }
}
