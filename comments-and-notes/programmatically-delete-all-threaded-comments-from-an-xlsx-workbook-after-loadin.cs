using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class DeleteThreadedComments
    {
        public static void Run()
        {
            Workbook workbook = new Workbook("input.xlsx");
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.ClearComments();
            }
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            DeleteThreadedComments.Run();
        }
    }
}