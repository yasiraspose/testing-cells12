using System;
using Aspose.Cells;

public class TransferWorksheet
{
    public static void Run()
    {
        Workbook sourceWorkbook = new Workbook("source.xlsx");
        Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

        Workbook targetWorkbook = new Workbook();
        Worksheet targetSheet = targetWorkbook.Worksheets[0];
        targetSheet.Name = "CopiedSheet";

        targetSheet.Copy(sourceSheet);

        targetWorkbook.Save("target.xlsx");
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        TransferWorksheet.Run();
    }
}