using System;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace C_Excel
{
    class XmlExcelBridge
    {
        static void Main(string[] args)
        {
            XDocument xmlInput = XmlMedian.XmlParser();
            string excelOutput = ExcelMedian.Finder();
            ExcelMedian.Writer(xmlInput, excelOutput);
        }
    }
}
