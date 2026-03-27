using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Path to the source Excel workbook
        string sourcePath = "input.xlsx";

        // Desired path for the generated PowerPoint file
        string destPath = "output.pptx";

        // Convert the Excel file to PPTX using Aspose.Cells ConversionUtility
        ConversionUtility.Convert(sourcePath, destPath);

        Console.WriteLine("Excel workbook has been successfully converted to PowerPoint presentation.");
    }
}