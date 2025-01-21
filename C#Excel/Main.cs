using System;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace C_Excel
{
    class XmlExcelBridge
    {
        /// <summary>
        /// Query users for XML and Excel files. Open Excel based off of required specifications.
        /// <para>Find values for each header name in XML and insert them under them under the corresponding headers.</para>
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            XDocument xmlInput = XmlMedian.Finder();
            string excelOutput = ExcelMedian.Finder();
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(excelOutput, true))
            {
                WorkbookPart workbookPart = doc.WorkbookPart;
                Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(sheet => sheet.Name == "GLOBAL");
                WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();


                // Grab First Row of GLOBAL sheet
                Row firstRow = sheetData.Elements<Row>().ElementAtOrDefault(0);
                if (firstRow != null)
                {
                    
                    foreach (Cell cellData in firstRow)
                    {
                        // Grab Value from current cell
                        string cellValue = ExcelMedian.GetCellValue(doc, cellData);

                        // If a cell's value == xml tag name
                        if (XmlMedian.FindTag(cellValue, xmlInput))
                        {
                            string xmlTagValue = XmlMedian.GetTagValue(cellValue, xmlInput);
                            // Go down in rows (same column) until an empty cell is found where data can be inserted
                            string emptyCellRef = ExcelMedian.FindEmptyCellRef(doc, cellData.CellReference, worksheetPart);
                            // Insert data into empty cell
                            ExcelMedian.InsertInCell(doc, emptyCellRef, xmlTagValue);

                        }

                    }
                }

            }
        }
    }
}
