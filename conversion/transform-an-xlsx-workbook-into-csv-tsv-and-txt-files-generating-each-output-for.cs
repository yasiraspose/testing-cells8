using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX workbook
            string sourcePath = "source.xlsx";

            // Load the workbook using the constructor that accepts a file name
            Workbook workbook = new Workbook(sourcePath);

            // Save as CSV (comma‑separated values)
            workbook.Save("output.csv", SaveFormat.Csv);

            // Save as TSV (tab‑separated values)
            workbook.Save("output.tsv", SaveFormat.Tsv);

            // Save as a generic TXT file.
            // TxtSaveOptions allows custom separators; here we use a tab character.
            TxtSaveOptions txtOptions = new TxtSaveOptions(SaveFormat.Csv);
            txtOptions.SeparatorString = "\t";   // Tab separator for TXT
            workbook.Save("output.txt", txtOptions);

            Console.WriteLine("Conversion completed: CSV, TSV, and TXT files generated.");
        }
    }
}