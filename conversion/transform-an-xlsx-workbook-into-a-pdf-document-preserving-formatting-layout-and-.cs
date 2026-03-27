using System;
using Aspose.Cells;

public class XlsxToPdfConverter
{
    // Converts an existing XLSX workbook to PDF while preserving formatting, layout, and pagination.
    public static void Convert(string sourcePath, string destinationPath)
    {
        // Load the workbook from the XLSX file.
        Workbook workbook = new Workbook(sourcePath);

        // Save the workbook as PDF. This retains all visual aspects of the original file.
        workbook.Save(destinationPath, SaveFormat.Pdf);

        Console.WriteLine($"Conversion completed: \"{sourcePath}\" → \"{destinationPath}\"");
    }

    // Entry point for the application.
    public static void Main(string[] args)
    {
        if (args.Length != 2)
        {
            Console.WriteLine("Usage: XlsxToPdfConverter <source.xlsx> <destination.pdf>");
            return;
        }

        string sourcePath = args[0];
        string destinationPath = args[1];

        try
        {
            Convert(sourcePath, destinationPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during conversion: {ex.Message}");
        }
    }
}