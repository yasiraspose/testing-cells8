using System;
using Aspose.Cells;

namespace AsposeCellsConversionDemo
{
    public class WorkbookConverter
    {
        public static void Run()
        {
            string sourcePath = "input.xlsx";
            string odsPath = "output.ods";
            string sxcPath = "output.sxc";
            string fodsPath = "output.fods";

            try
            {
                Workbook wb = new Workbook(sourcePath);

                wb.Save(odsPath, SaveFormat.ODS);
                Console.WriteLine($"Successfully converted to ODS: {odsPath}");

                wb.Save(sxcPath, SaveFormat.SXC);
                Console.WriteLine($"Successfully converted to SXC: {sxcPath}");

                wb.Save(fodsPath, SaveFormat.FODS);
                Console.WriteLine($"Successfully converted to FODS: {fodsPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Conversion failed: {ex.Message}");
            }
        }

        public static void Main()
        {
            Run();
        }
    }
}