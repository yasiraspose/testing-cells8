using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsXpsConversion
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX workbook
            string sourcePath = "input.xlsx";

            // Desired output XPS file path
            string outputPath = "output.xps";

            // Convert the Excel workbook to XPS format.
            // This uses the built‑in ConversionUtility which preserves data fidelity and layout.
            ConversionUtility.Convert(sourcePath, outputPath);

            Console.WriteLine($"Conversion completed: '{sourcePath}' -> '{outputPath}'");
        }
    }
}