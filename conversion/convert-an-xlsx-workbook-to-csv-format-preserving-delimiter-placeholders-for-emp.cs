using System;
using System.Text;
using Aspose.Cells;

class ConvertXlsxToCsv
{
    static void Main()
    {
        // Source XLSX file and destination CSV file paths
        string sourcePath = "input.xlsx";
        string destPath = "output.csv";

        // Load the workbook from the XLSX file
        Workbook workbook = new Workbook(sourcePath);

        // Configure CSV save options:
        // - KeepSeparatorsForBlankRow = true ensures delimiters are written for completely blank rows
        // - TrimLeadingBlankRowAndColumn = false preserves leading blank rows/columns
        // - TrimTailingBlankCells = false keeps trailing empty cells in each row
        TxtSaveOptions csvOptions = new TxtSaveOptions
        {
            Encoding = Encoding.UTF8,
            KeepSeparatorsForBlankRow = true,
            TrimLeadingBlankRowAndColumn = false,
            TrimTailingBlankCells = false
        };

        // Save the workbook as CSV using the configured options
        workbook.Save(destPath, csvOptions);
    }
}