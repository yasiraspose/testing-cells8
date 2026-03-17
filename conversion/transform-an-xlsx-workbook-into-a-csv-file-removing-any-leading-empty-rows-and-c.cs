using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source XLSX workbook
        string sourcePath = "input.xlsx";

        // Load the workbook from the file
        Workbook workbook = new Workbook(sourcePath);

        // Configure CSV save options to trim leading empty rows and columns
        TxtSaveOptions csvOptions = new TxtSaveOptions
        {
            TrimLeadingBlankRowAndColumn = true // Remove leading blanks like Excel does
        };

        // Path for the resulting CSV file
        string destinationPath = "output.csv";

        // Save the workbook as CSV using the configured options
        workbook.Save(destinationPath, csvOptions);
    }
}