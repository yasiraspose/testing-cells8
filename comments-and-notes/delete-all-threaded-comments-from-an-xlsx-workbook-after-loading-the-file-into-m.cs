using System;
using Aspose.Cells;

namespace AsposeCellsThreadedCommentRemoval
{
    class Program
    {
        static void Main()
        {
            // Load the existing XLSX workbook into memory
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets and clear all comments (including threaded comments)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // This method removes every comment from the worksheet
                sheet.ClearComments();
            }

            // Save the workbook after removing the comments
            string outputPath = "output_without_threaded_comments.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
    }
}