using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace MinimalExcel
{
    class ExcelControl
    {
        ExcelControl()
        {
            // COM Objekte deklarieren und leer initialisieren
            Excel.Application   excelApp = null;
            Excel.Workbook      excelWorkbook = null;
            Excel.Workbooks     excelWorkbooks = null;
            Excel.Worksheet     excelSheet = null;
            Excel.Range         excelBereich = null;

            // Objekt erzeugen
            Console.WriteLine("Erzeuge Excel COM Objekt");
            excelApp = new Excel.Application();
            excelApp.Visible = true;

            // Dateipfad der xlsx zusammenbauen
            Console.WriteLine("Öffne Datei");
            String filename = "ExcelFile.xlsx"; // die Excel Datei sollte im Projekt liegen und immer in das Ausgabeverzeichnis kopiert werden
            String path = System.IO.Path.GetFullPath(filename);
            if (System.IO.File.Exists(path))
            {
                // Öffnen einer Excel Datei
                excelWorkbooks = excelApp.Workbooks;
                excelWorkbook = excelWorkbooks.Open(path);
                excelSheet = (Excel.Worksheet)excelApp.ActiveSheet;

                // Lesen der Datei
                Console.WriteLine("Lese aus der Datei");
                excelBereich = excelSheet.Cells[1, "A"] as Excel.Range;
                String zellWert = (String)excelBereich.Value;
                Console.WriteLine("   Wert: " + zellWert);

                Excel.Range workSheet_range = excelSheet.get_Range("A1", "B3"); // Mehrere Zellen erhelten Sie auch so

                // Schreiben in die Datei
                Console.WriteLine("Schreibe in die Datei");
                excelSheet.Cells[1, "B"] = "Hallo Welt";
                excelSheet.Cells[2, "B"] = "1";
                excelSheet.Cells[3, "B"] = "=A2+2";

                // Speichern der Datei und beenden.
                Console.WriteLine("Speichern der Änderungen");
                String newFileName = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(path), "newExcelFile.xlsx");
                excelSheet.SaveAs(newFileName);
            }


            // Alle Excel COM Objekte der Reihe nach schließen bzw freigeben 
            if (excelBereich != null)   Marshal.ReleaseComObject(excelBereich);
            if (excelSheet != null)     Marshal.ReleaseComObject(excelSheet);
            excelWorkbook.Close(0);
            if (excelWorkbook != null)  Marshal.ReleaseComObject(excelWorkbook);
            excelWorkbooks.Close();
            if (excelWorkbooks != null) Marshal.ReleaseComObject(excelWorkbooks);
            excelApp.Quit();
            if (excelApp != null)       Marshal.ReleaseComObject(excelApp);

            // Console.ReadKey();
        }


        static void Main(string[] args)
        {
            new ExcelControl();
        }
    }
}
