using System;
using System.Reflection.Metadata.Ecma335;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Spreadsheet;



namespace C_Excel
{
    /**
     * @author Derek Smith
     * 
     */
    public class XmlMedian
    {
        /// <summary>
        /// Receives user input for path to XML file. Validates whether file exists and is an XML file.
        /// </summary>
        /// <returns>XDocument variable representing an loaded XML file.</returns>
        public static XDocument Finder()
        {
            Console.WriteLine("Input Path to XML File");
            string xmlInputPath = Console.ReadLine();

            XDocument xmlDoc = XDocument.Load(xmlInputPath);
            while ((!File.Exists(xmlInputPath)) || (Path.GetExtension(xmlInputPath) != ".xml"))
            {
                Console.WriteLine("File cannot be found or file is not an XML file.\n Input new filepath: ");
                xmlInputPath = Console.ReadLine();
            }

            Console.WriteLine(xmlDoc.ToString());

            return xmlDoc;
        }

        /// <summary>
        /// Looks for a tag by name in an XML document and retrieves the (value of the first tag that applies).
        /// <para>Replaces all whitespace in input with underscores (_) to avoid XML tag error.</para>
        /// </summary>
        /// <param name="tagName">string variable represeting the tag name that will be searched for.</param>
        /// <param name="xmlDoc">XDocument variable representing a loaded XML file.</param>
        /// <returns>string variable representing the value contained in a specific tag.</returns>
        public static string GetTagValue(string tagName, XDocument xmlDoc) 
        {
            string tagValue;
            tagName = tagName.Replace(" ", "_");

            // null should not be returned as FindTag should be called prior to this
            var tag = xmlDoc.Descendants(tagName).FirstOrDefault();

            tagValue = tag?.Value;
            return tagValue;
        }

        /// <summary>
        /// Verifies if tag exists in XML document.
        /// <para>Replaces all whitespace in input with underscores (_) to avoid XML tag error.</para>
        /// </summary>
        /// <param name="tagName">string variable representing the tag name that is being searched for.</param>
        /// <param name="xmlDoc">XDocument variable representing a loaded XML file.</param>
        /// <returns>bool with true value representing that the tag exists and false representing the opposite.</returns>
        public static bool FindTag(string tagName, XDocument xmlDoc)
        {
            tagName = tagName.Replace(" ", "_");

            // First character being a digit invalidates it as being an XML tag
            // check exists to prevent XML error
            if (!Char.IsLetter(tagName[0]))
            {
                return false;
            }

            // if tag with name input is not found return false
            if (xmlDoc.Descendants(tagName).FirstOrDefault() == null)
            {
                return false;
            }
            
            return true;
        }
    }
}





















