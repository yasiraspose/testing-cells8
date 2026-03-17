using System;
using Aspose.Cells;

class InsertCommentDemo
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (you can change the index as needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a comment to cell A1 using the Comments collection
        int commentIndex = worksheet.Comments.Add("A1");
        Comment comment = worksheet.Comments[commentIndex];

        // Set the comment text and optional properties
        comment.Note = "This is a new comment added via code.";
        comment.Author = "AsposeUser";
        comment.IsVisible = true; // Make the comment visible by default

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}