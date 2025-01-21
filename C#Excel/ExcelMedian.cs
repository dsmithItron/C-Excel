using System;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;

namespace C_Excel
{
    /**
     * @author Derek Smith
     * 
     */
    class ExcelMedian
    {
        /// <summary>
        /// Requests a filepath from a user. Validates the filepath until it leads to a Excel Spreadsheet file.
        /// </summary>
        /// <returns> A string representing a path to an Excel Spreadsheet.</returns>
        public static string Finder()
        {
            Console.WriteLine("Input path to Excel File");

            string filePath = Console.ReadLine();
            // Continue prompting for file if the file does not exist or the file is not an Excel spreadsheet
            while ((!File.Exists(filePath)) || (Path.GetExtension(filePath) != ".xlsx"))
            {
                Console.WriteLine("File cannot be found or file is not an Excel file.\n Input new filepath: ");
                filePath = Console.ReadLine();
            }

            return filePath;
        }

        /// <summary>
        /// Uses a spreadsheet document and cell attached to that document to obtain the value of the cell.
        /// <para>Verifies that the cell exists and if the cell is a Shared String pulls the cell's value from the Shared Table.</para>
        /// <para>Returns a string representing the cell's value.</para>
        /// </summary>
        /// <param name="doc"> SpreadsheetDocument variable representing an open Spreadsheet that is being viewed.</param>
        /// <param name="cell"> Cell variable that represents a cell in the spreadsheet.</param>
        /// <returns>Returns a string representing the cell's value.</returns>
        public static string GetCellValue(SpreadsheetDocument doc, Cell cell)
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

        /// <summary>
        /// Uses a SpreadsheetDocument and cell reference to insert a value into a specific cell.
        /// </summary>
        /// <param name="doc">SpreadsheetDocument variable representing an open Spreadsheet that is being edited.</param>
        /// <param name="cellRef">string representing a cell reference. Format Ex. = (A3, B4, BB4, A44)</param>
        /// <param name="insertValue">string representing text that will be inserted into a cell.</param>
        /// <returns>bool with true representing a successful insert with false representing the opposite.</returns>
        public static bool InsertInCell(SpreadsheetDocument doc, string cellRef, string insertValue)
        {
            string column = SplitCellRef(cellRef)[0];
            // converts Row value from int to uint in order to use Microsoft function (InsertCellInWorksheet)
            uint uRow = Convert.ToUInt32(SplitCellRef(cellRef)[1]);


            Cell insertCell = InsertCellInWorksheet(doc, column, uRow);
            insertCell.CellValue = new CellValue(insertValue);
            insertCell.DataType = new EnumValue<CellValues>(CellValues.String);
            if (GetCellValue(doc, insertCell) == insertValue)
            {
                return true;
            }
            return false;
        }

        /// <summary>
        /// Goes down 1 row at a time from starting cell until empty cell is found.
        /// </summary>
        /// <param name="doc">SpreadsheetDocument variable representing an open Spreadsheet that is being viewed.</param>
        /// <param name="cellData">Cell variable representing the cell</param>
        /// <param name="worksheetPart"></param>
        /// <returns>String representing a reference to an empty cell. Format Ex. = (A3, B4, BB4, A44)</returns>
        public static string FindEmptyCellRef(SpreadsheetDocument doc, string cellRef, WorksheetPart worksheetPart)
        {
            string cellValue = "empty";
            
            // obtains column portion of cellRef
            string column = SplitCellRef(cellRef)[0];
            // obtains row portion of cellRef and subtracts to deal with initial addition in loop
            int row = Int32.Parse(SplitCellRef(cellRef)[1]) - 1;

            // obtains Cell where cellReference = cellRef
            Cell cellData = worksheetPart.Worksheet.Descendants<Cell>().Where(c => c.CellReference == cellRef).FirstOrDefault();

            while (cellValue != "")
            {
                row++;
                // CellRef is pushed down a row
                cellRef = column + row;
                cellData = worksheetPart.Worksheet.Descendants<Cell>().Where(c => c.CellReference == cellRef).FirstOrDefault();
                // solving for the cell not existing
                if (cellData == null || (GetCellValue(doc, cellData) == null))
                {
                    cellValue = "";
                }
                else
                {
                    cellValue = GetCellValue(doc, cellData);
                }
            }

            return cellRef;
        }

        /// <summary>
        /// Splits CellRef into two substrings containing the column and row of a cell.
        /// </summary>
        /// <param name="inputString">string representing a cell reference.</param>
        /// <returns> List containing the cell column at index 0 and the cell row at index 1.</returns>
        public static List<string> SplitCellRef(string inputString)
        {
            List<string> subStrings = [];
            int dividePoint = 0;

            foreach (char c in inputString)
            {
                // if current character is not a digit
                if (!char.IsDigit(c))
                {
                    dividePoint++;
                }
                else
                {
                    break; // Stop once we encounter a non-digit character
                }
            }

            // Adding column substring
            subStrings.Add(inputString.Substring(0, dividePoint));
            // Adding row substring
            subStrings.Add(inputString.Substring(dividePoint));


            return subStrings;
        }
        /// <summary>
        /// Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="columnName"></param>
        /// <param name="rowIndex"></param>
        /// <returns>Returns the cell if it already exists</returns>
        public static Cell InsertCellInWorksheet(SpreadsheetDocument doc, string columnName, uint rowIndex)
        {
            WorkbookPart workbookPart = doc.WorkbookPart;
            Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(sheet => sheet.Name == "GLOBAL");
            WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
            string cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row;

            if (sheetData?.Elements<Row>().Where(r => r.RowIndex is not null && r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData!.Elements<Row>().Where(r => r.RowIndex is not null && r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // If there is not a cell with the specified column name, insert one.  
            if (row.Elements<Cell>().Where(c => c.CellReference is not null && c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference is not null && c.CellReference.Value == cellReference).First();
            }
            else
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                Cell? refCell = null;

                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (string.Compare(cell.CellReference?.Value, cellReference, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                return newCell;
            }
        }
    }
}
