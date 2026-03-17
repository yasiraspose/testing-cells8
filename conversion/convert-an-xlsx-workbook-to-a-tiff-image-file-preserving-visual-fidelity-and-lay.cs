using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

class ConvertXlsxToTiff
{
    static void Main()
    {
        // Path to the source XLSX file
        string inputPath = "input.xlsx";

        // Desired output TIFF file path
        string outputPath = "output.tiff";

        // Load the workbook from the XLSX file
        Workbook workbook = new Workbook(inputPath);

        // Configure image rendering options for TIFF output
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            ImageType = ImageType.Tiff,                     // Set output format to TIFF
            TiffCompression = TiffCompression.CompressionLZW, // Use LZW compression for better quality
            HorizontalResolution = 300,                     // Set horizontal DPI
            VerticalResolution = 300,                       // Set vertical DPI
            OnePagePerSheet = true                          // Render each worksheet as a separate page
        };

        // Create a renderer for the entire workbook
        WorkbookRender renderer = new WorkbookRender(workbook, options);

        // Render the workbook to a multi‑page TIFF file
        renderer.ToImage(outputPath);

        Console.WriteLine($"Workbook successfully converted to TIFF: {outputPath}");
    }
}