using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace openXml3
{
    class Program
    {
        static void Main(string[] args)
        {
            //CreateSpreadsheetWorkbook("C:\\Temp\\Test2.xlsx");
            //SalesReportBuilder s = new SalesReportBuilder();
            //s.CreateDocument();
            jdoc j = new jdoc();
            j.createDoc();
        }

        public static void CreateSpreadsheetWorkbook(string filepath)
        {

            go2();
            return;
            Console.ReadLine();
            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.

            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook);

            // Add a WorkbookPart to the document.
            WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook.
            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

            // Append a new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet the bed" };
            sheets.Append(sheet);

            sheet = new Sheet() { Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 2, Name = "Clean the bed" };
            sheets.Append(sheet);
            workbookpart.Workbook.Save();

            // Close the document.
            spreadsheetDocument.Close();
        }

        public static void go2()
        {
            int[,] demo = giveArray();

            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Create(
            System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "c:\\temp\\arse.xlsx"),
            SpreadsheetDocumentType.Workbook))
            {
                // create the workbook
                spreadSheet.AddWorkbookPart();
                spreadSheet.WorkbookPart.Workbook = new Workbook();     // create the worksheet
                spreadSheet.WorkbookPart.AddNewPart<WorksheetPart>();
                spreadSheet.WorkbookPart.WorksheetParts.First().Worksheet = new Worksheet();

                // create sheet data
                spreadSheet.WorkbookPart.WorksheetParts.First().Worksheet.AppendChild(new SheetData());

                for (int i = 0; i < 10; i++ )
                {
                    // create row
                    spreadSheet.WorkbookPart.WorksheetParts.First().Worksheet.First().AppendChild(new Row());

                    for(int n = 0; n < 10; n++)
                    {
                        // create cell with data
                        spreadSheet.WorkbookPart.WorksheetParts.First().Worksheet.First().Last().AppendChild(
                              new Cell() { CellValue = new CellValue(demo[i,n].ToString()) });
                    }


                }

                    



                //spreadSheet.WorkbookPart.WorksheetParts.First().Worksheet.First().First().AppendChild(
                //        new Cell() { CellValue = new CellValue("102") });

                // save worksheet
                spreadSheet.WorkbookPart.WorksheetParts.First().Worksheet.Save();

                // create the worksheet to workbook relation
                spreadSheet.WorkbookPart.Workbook.AppendChild(new Sheets());
                spreadSheet.WorkbookPart.Workbook.GetFirstChild<Sheets>().AppendChild(new Sheet()
                {
                    Id = spreadSheet.WorkbookPart.GetIdOfPart(spreadSheet.WorkbookPart.WorksheetParts.First()),
                    SheetId = 1,
                    Name = "test"
                });

                spreadSheet.WorkbookPart.Workbook.Save();
            }
        }

        public static int[,] giveArray()
        {
            int[,] ret = new int[10,10];

            for (int i = 0; i < 10; i++)
            {
                for (int n = 0; n < 10; n++)
                {
                    ret[i, n] = i * n;
                }
            }
            return ret;
        }


    }
}
