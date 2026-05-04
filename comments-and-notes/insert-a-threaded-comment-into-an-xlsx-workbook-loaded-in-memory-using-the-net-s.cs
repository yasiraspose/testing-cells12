using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsThreadedCommentDemo
{
    public class ThreadedCommentHelper
    {
        // Loads an existing workbook from a memory stream, adds a threaded comment,
        // and returns the modified workbook as a new memory stream.
        public static MemoryStream AddThreadedCommentToWorkbook(MemoryStream inputWorkbookStream)
        {
            // Load the workbook from the provided memory stream (in‑memory workbook)
            Workbook workbook = new Workbook(inputWorkbookStream);

            // Get the first worksheet (or any target worksheet)
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a threaded comment author (name, userId, provider)
            int authorIndex = workbook.Worksheets.ThreadedCommentAuthors.Add(
                "John Doe",          // Author name
                "john.doe@example.com", // User ID / email
                "EXAMPLE_PROVIDER");    // Provider identifier

            ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors[authorIndex];

            // Add a threaded comment to cell B2 (row 1, column 1)
            // Parameters: row index, column index, comment text, author object
            worksheet.Comments.AddThreadedComment(1, 1, "This is a threaded comment added via code.", author);

            // Optionally, retrieve the comment collection for verification
            ThreadedCommentCollection comments = worksheet.Comments.GetThreadedComments(1, 1);
            foreach (ThreadedComment tc in comments)
            {
                Console.WriteLine($"Comment by {tc.Author.Name}: {tc.Notes}");
            }

            // Save the modified workbook to a new memory stream
            MemoryStream outputStream = new MemoryStream();
            workbook.Save(outputStream, SaveFormat.Xlsx);

            // Reset the position of the stream so it can be read from the beginning
            outputStream.Position = 0;

            return outputStream;
        }

        // Example usage
        public static void Main()
        {
            // Create a simple workbook in memory to demonstrate loading
            Workbook originalWorkbook = new Workbook();
            originalWorkbook.Worksheets[0].Cells["A1"].PutValue("Sample Data");

            // Save the original workbook to a memory stream
            MemoryStream originalStream = new MemoryStream();
            originalWorkbook.Save(originalStream, SaveFormat.Xlsx);
            originalStream.Position = 0;

            // Add a threaded comment to the workbook loaded from memory
            MemoryStream modifiedStream = AddThreadedCommentToWorkbook(originalStream);

            // For demonstration, write the result to a file (optional)
            using (FileStream file = new FileStream("ThreadedCommentResult.xlsx", FileMode.Create, FileAccess.Write))
            {
                modifiedStream.CopyTo(file);
            }

            Console.WriteLine("Threaded comment added and workbook saved as 'ThreadedCommentResult.xlsx'.");
        }
    }
}