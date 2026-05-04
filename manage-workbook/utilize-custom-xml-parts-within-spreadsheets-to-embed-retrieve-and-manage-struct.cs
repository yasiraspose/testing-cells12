using System;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsCustomXmlDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Define XML data and an optional XML schema
            string xmlData = "<Employee><Name>John Doe</Name><Age>30</Age></Employee>";
            string xmlSchema = "<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>" +
                               "<xs:element name='Employee'>" +
                               "<xs:complexType>" +
                               "<xs:sequence>" +
                               "<xs:element name='Name' type='xs:string'/>" +
                               "<xs:element name='Age' type='xs:int'/>" +
                               "</xs:sequence>" +
                               "</xs:complexType>" +
                               "</xs:element>" +
                               "</xs:schema>";

            // Convert strings to UTF‑8 byte arrays
            byte[] dataBytes = Encoding.UTF8.GetBytes(xmlData);
            byte[] schemaBytes = Encoding.UTF8.GetBytes(xmlSchema);

            // Add a custom XML part to the workbook
            int partIndex = wb.CustomXmlParts.Add(dataBytes, schemaBytes);

            // Save the workbook containing the custom XML part
            string filePath = "CustomXmlDemo.xlsx";
            wb.Save(filePath);

            // Load the workbook back
            Workbook loadedWb = new Workbook(filePath);

            // Retrieve the custom XML part by its index
            var retrievedPart = loadedWb.CustomXmlParts[partIndex];

            if (retrievedPart != null)
            {
                string retrievedXml = Encoding.UTF8.GetString(retrievedPart.Data);
                Console.WriteLine("Retrieved XML:");
                Console.WriteLine(retrievedXml);

                // Update the XML content of the retrieved part
                string updatedXml = "<Employee><Name>Jane Smith</Name><Age>28</Age></Employee>";
                retrievedPart.Data = Encoding.UTF8.GetBytes(updatedXml);
            }

            // Save the workbook with the updated custom XML part
            string updatedFilePath = "CustomXmlDemo_Updated.xlsx";
            loadedWb.Save(updatedFilePath);
        }
    }
}