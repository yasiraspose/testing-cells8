using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsMhtmlConversion
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel workbook (XLSX)
            string sourcePath = "input.xlsx";

            // Destination path for the MHTML file.
            // The .mht extension signals MHTML format (SaveFormat.MHtml).
            string destPath = "output.mht";

            // Convert the Excel file to MHTML while preserving all data and formatting.
            // This uses the provided ConversionUtility.Convert(string, string) rule.
            ConversionUtility.Convert(sourcePath, destPath);

            Console.WriteLine($"Conversion completed: '{sourcePath}' -> '{destPath}'");
        }
    }
}