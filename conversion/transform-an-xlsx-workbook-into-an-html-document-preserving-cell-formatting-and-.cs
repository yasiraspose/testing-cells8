using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source XLSX workbook
        string sourcePath = "input.xlsx";

        // Desired path for the generated HTML file
        string htmlPath = "output.html";

        // Load the workbook from the XLSX file (lifecycle: create/load)
        Workbook workbook = new Workbook(sourcePath);

        // Configure HTML save options to retain formatting and content
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.ExportActiveWorksheetOnly = false;   // export all worksheets
        saveOptions.ExportGridLines = true;              // keep grid lines visible
        saveOptions.ExportWorkbookProperties = true;    // include workbook properties
        saveOptions.ExportWorksheetProperties = true;   // include worksheet properties
        saveOptions.CalculateFormula = true;            // evaluate formulas before export
        saveOptions.HtmlVersion = Aspose.Cells.HtmlVersion.Html5; // use modern HTML5

        // Save the workbook as an HTML document (lifecycle: save)
        workbook.Save(htmlPath, saveOptions);

        Console.WriteLine("Workbook successfully converted to HTML.");
    }
}