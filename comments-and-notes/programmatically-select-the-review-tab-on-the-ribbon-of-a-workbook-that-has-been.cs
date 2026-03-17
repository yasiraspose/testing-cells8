using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing workbook from an XLSX file
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Define Ribbon XML that selects the built‑in Review tab when the workbook is opened
        string ribbonXml =
            "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
            "  <ribbon>" +
            "    <tabs>" +
            "      <tab idMso=\"TabReview\"/>" +   // Built‑in Review tab
            "    </tabs>" +
            "  </ribbon>" +
            "</customUI>";

        // Apply the custom Ribbon UI to the workbook
        workbook.RibbonXml = ribbonXml;

        // Save the modified workbook (preserving the original format)
        string outputPath = "output.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}