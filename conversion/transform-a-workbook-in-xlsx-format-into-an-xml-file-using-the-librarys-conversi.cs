using System;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Path to the source XLSX workbook
        string sourcePath = "input.xlsx";

        // Desired output XML file path
        string destPath = "output.xml";

        // Convert the Excel workbook to XML using Aspose.Cells ConversionUtility
        ConversionUtility.Convert(sourcePath, destPath);

        Console.WriteLine("Workbook has been successfully converted to XML.");
    }
}