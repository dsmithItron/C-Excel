using System;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;

namespace C_Excel
{
    class ExcelMedian
    {
        public static string Finder()
        {
            Console.WriteLine("Input Excel File");

            string filePath = Console.ReadLine();
            while ((!File.Exists(filePath)) || (filePath.ToLower() == "stop") || (Path.GetExtension(filePath) != ".xlsx"))
            {
                Console.WriteLine("File cannot be found or file is not an Excel file.\n Input new filepath or \"stop\" \n");
                filePath = Console.ReadLine();
            }

            return filePath;
        }

        public static void Writer(XDocument xmlInput, string excelPath)
        {
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(excelPath, true))
            {
                WorkbookPart workbookPart = doc.WorkbookPart;
                Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(sheet => sheet.Name == "Global");
                WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();


                // for each cell in the first row
                // if cell value == any tag name in XML 
                // drop rows till cell value is empty
                // input value in XML tag
                Row firstRow = sheetData.Elements<Row>().ElementAtOrDefault(0);
                foreach (Cell celldata in firstRow)
                {
                    string cellValue = ExcelMedian.GetCellValue(doc, celldata);

                    // If a CellValue == An XML tag name
                    if (XmlMedian.FindTag(cellValue, xmlInput))
                    {
                        string xmlTagValue = XmlMedian.GetTagValue(cellValue, xmlInput);
                        // Drop rows until we encounter an empty cell
                        foreach (Row row in sheetData.Elements<Row>())
                        {
                            if (row.Elements<Cell>().Any(c => GetCellValue(doc, celldata) == ""))
                            {
                                // Stop processing once an empty cell is found
                                break;
                            }
                            if(InsertInCell());
                        }
                    }

                }

            }


        }

        static string GetCellValue(SpreadsheetDocument doc, Cell cell)
        {
            string value = string.Empty;


            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                // Extract the index of the shared string
                int index = int.Parse(cell.CellValue.Text);

                // Get the shared string table part
                SharedStringTablePart sharedStringTablePart = doc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                // Get the actual shared string value using the index
                value = sharedStringTablePart.SharedStringTable.ElementAt(index).InnerText;
            }

            else if (cell.CellValue != null)
            {
                value = cell.CellValue.Text;
            }

            return value;
        }
        /*
        public static void Median()
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
        */

        // does this need overload for different inserts?
        // Maybe just give it the whole xml and let it figure it out?
        static bool InsertInCell(int CellRef, string insertValue)
        {
            return false;
        }
    }
}
