using System;
using System.Reflection.Metadata.Ecma335;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Spreadsheet;



namespace C_Excel
{
    public class XmlMedian
    {
        public static XDocument xmlParser()
        {
            Console.WriteLine("Input Path to XML File");
            string xmlInputPath = Console.ReadLine();

            XDocument xmlDoc = XDocument.Load(xmlInputPath);

            Console.WriteLine(xmlDoc.ToString());

            return xmlDoc;
        }

        public static string GetTagValue(string tagName, XDocument xmlDoc) 
        {
            string tagValue;

            var tag = xmlDoc.Descendants(tagName).FirstOrDefault();
            tagValue = tag?.Value;
            return tagValue;
        }

        public static bool FindTag(string tagName, XDocument xmlDoc)
        {
            if (xmlDoc.Descendants(tagName).FirstOrDefault() == null) {
                return false;
            }
            return true;
        }
    }
}





















