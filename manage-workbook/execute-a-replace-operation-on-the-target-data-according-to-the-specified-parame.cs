using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing.Texts;

class ReplaceWithFormattingDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Put initial text into cell A1
        Cell cell = sheet.Cells["A1"];
        cell.PutValue("Hello World");

        // Apply formatting to the word "World" (start index 6, length 5)
        cell.Characters(6, 5).Font.IsBold = true;
        cell.Characters(6, 5).Font.Color = Color.Blue;
        cell.Characters(6, 5).Font.Size = 14;

        // Prepare ReplaceOptions with font settings for the replacement text
        ReplaceOptions options = new ReplaceOptions
        {
            CaseSensitive = true,
            MatchEntireCellContents = false
        };

        // Define font settings for the new text "Universe"
        FontSetting replacementSetting = new FontSetting(6, 8, workbook.Worksheets);
        replacementSetting.Font.IsBold = true;
        replacementSetting.Font.Color = Color.Blue;
        replacementSetting.Font.Size = 14;

        options.FontSettings = new FontSetting[] { replacementSetting };

        // Perform the replacement
        workbook.Replace("World", "Universe", options);

        // Save the workbook
        workbook.Save("ReplaceWithFormattingDemo.xlsx");
    }
}