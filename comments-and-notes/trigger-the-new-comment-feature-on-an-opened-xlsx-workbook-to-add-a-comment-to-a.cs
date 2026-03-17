using System;
using Aspose.Cells;

namespace AsposeCellsCommentDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Get the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a comment to cell B2 using the CommentCollection.Add(string) method
            int commentIndex = worksheet.Comments.Add("B2");
            Comment comment = worksheet.Comments[commentIndex];

            // Set comment properties
            comment.Note = "This is a new comment added via Aspose.Cells.";
            comment.Author = "DemoUser";

            // Save the workbook with the newly added comment
            workbook.Save("output.xlsx");
        }
    }
}