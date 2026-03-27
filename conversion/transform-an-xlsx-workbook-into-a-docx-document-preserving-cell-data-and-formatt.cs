using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class Program
    {
        static void Main()
        {
            XlsxToDocxConverter.Run();
        }
    }

    public class XlsxToDocxConverter
    {
        public static void Run()
        {
            string sourcePath = "input.xlsx";
            string destPath = "output.docx";

            Workbook workbook = new Workbook(sourcePath);

            DocxSaveOptions saveOptions = new DocxSaveOptions();
            // The property SaveAsEditableShaps is obsolete; the default behavior preserves shapes.
            // If needed, use the newer API as per documentation.

            workbook.Save(destPath, saveOptions);

            Console.WriteLine("Excel workbook has been successfully converted to DOCX.");
        }
    }
}