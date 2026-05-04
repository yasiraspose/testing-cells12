using System;
using Aspose.Cells;

namespace ThreadedCommentTimestampDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the existing workbook
            string inputPath = "input.xlsx";

            // Load the workbook (uses the provided Workbook(string) constructor)
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the collection of comments on the current worksheet
                CommentCollection comments = sheet.Comments;

                // Loop through each comment in the collection
                for (int c = 0; c < comments.Count; c++)
                {
                    // Retrieve the comment object
                    Comment comment = comments[c];

                    // Obtain the cell address of the comment (row/column -> A1 style)
                    string cellAddress = CellsHelper.CellIndexToName(comment.Row, comment.Column);

                    // Access the threaded comments associated with this comment
                    ThreadedCommentCollection threadedComments = comment.ThreadedComments;

                    // Loop through each threaded comment
                    for (int t = 0; t < threadedComments.Count; t++)
                    {
                        ThreadedComment tc = threadedComments[t];

                        // Retrieve the creation timestamp using the CreatedTime property
                        DateTime createdTime = tc.CreatedTime;

                        // Output the information
                        Console.WriteLine($"Worksheet: {sheet.Name}, Cell: {cellAddress}, ThreadedComment #{t + 1}, CreatedTime: {createdTime}");
                    }
                }
            }

            // No need to save the workbook as we only read metadata
        }
    }
}