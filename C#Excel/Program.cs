using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace C_Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Input Excel File");

            string filePath = Console.ReadLine();
            if (filePath == string.Empty) { filePath = "sample.xlsx"; }
            if (!File.Exists(filePath))
            {
                using (SpreadsheetDocument doc = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
                {
                    WorkbookPart workbookPart = doc.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();
                    WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet(new SheetData());
                    Sheets sheets = doc.WorkbookPart.Workbook.AppendChild(new Sheets());
                    Sheet sheet = new Sheet() {
                        Id = doc.WorkbookPart.GetIdOfPart(worksheetPart), 
                        SheetId = 1, 
                        Name = "Sheet1" 
                    };
                    sheets.Append(sheet);
                    workbookPart.Workbook.Save();
                }
            }

            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = doc.WorkbookPart;
                Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().First();
                WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                if (!sheetData.Elements<Row>().Any())
                {
                    Console.WriteLine("Current Excel file is empty");
                }
                else
                {
                    Console.WriteLine("This is the current content of your file:");
                    foreach (Row row in sheetData.Elements<Row>())
                    {
                        foreach (Cell cell in row.Elements<Cell>())
                        {
                            Console.Write(cell.CellValue.Text + "\t");
                        }
                    }
                }
            }

            Console.WriteLine("\nWould you like to add a new row? (y/n)");
            string response = Console.ReadLine();

            if(response.ToLower() == "y")
            {
                Console.WriteLine("Enter data seperated by commas");
                string[] rowData = Console.ReadLine().Split(',');

                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(filePath, true))
                {
                    WorkbookPart workbookPart = doc.WorkbookPart;
                    Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().First();
                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                    Row newRow = new Row();
                    foreach (string cellData in rowData)
                    {
                        Cell cell = new Cell()
                        {
                            DataType = CellValues.String,
                            CellValue = new CellValue(cellData)
                        };
                        newRow.Append(cell);
                    }
                    sheetData.Append(newRow);
                }

            }

        }
    }
}
