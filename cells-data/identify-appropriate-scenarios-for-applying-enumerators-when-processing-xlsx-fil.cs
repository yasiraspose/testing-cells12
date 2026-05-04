using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsEnumeratorScenarios
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("SampleData.xlsx");
            Worksheet sheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // Scenario 1: Iterate over every populated cell in the worksheet.
            // ------------------------------------------------------------
            IEnumerator cellEnum = sheet.Cells.GetEnumerator();
            Console.WriteLine("All cells with values:");
            while (cellEnum.MoveNext())
            {
                Cell cell = (Cell)cellEnum.Current;
                if (cell.Value != null)
                {
                    Console.WriteLine($"{cell.Name}: {cell.Value}");
                }
            }

            // ------------------------------------------------------------
            // Scenario 2: Iterate over rows to process row‑level information
            // ------------------------------------------------------------
            IEnumerator rowEnum = sheet.Cells.Rows.GetEnumerator();
            Console.WriteLine("\nRow information:");
            while (rowEnum.MoveNext())
            {
                Row row = (Row)rowEnum.Current;
                double sum = 0;
                for (int col = 0; col < 3; col++)
                {
                    Cell c = row[col];
                    if (c != null && c.Type == CellValueType.IsNumeric)
                    {
                        sum += c.DoubleValue;
                    }
                }
                Console.WriteLine($"Row {row.Index} sum of first 3 columns = {sum}");
            }

            // ------------------------------------------------------------
            // Scenario 3: Iterate over a specific range.
            // ------------------------------------------------------------
            Aspose.Cells.Range dataRange = sheet.Cells.CreateRange("B2:D5");
            IEnumerator rangeEnum = dataRange.GetEnumerator();
            Console.WriteLine("\nCells in range B2:D5:");
            while (rangeEnum.MoveNext())
            {
                Cell cell = (Cell)rangeEnum.Current;
                Console.WriteLine($"{cell.Name}: {cell.Value}");
            }

            // ------------------------------------------------------------
            // Scenario 4: Iterate over pivot table row fields.
            // ------------------------------------------------------------
            if (sheet.PivotTables.Count == 0)
            {
                int pivotIdx = sheet.PivotTables.Add(dataRange.RefersTo, "F2", "PivotDemo");
                PivotTable pt = sheet.PivotTables[pivotIdx];
                pt.AddFieldToArea(PivotFieldType.Row, 0);   // First column as row field
                pt.AddFieldToArea(PivotFieldType.Data, 1); // Second column as data field
                pt.RefreshData();
                pt.CalculateData();
            }

            PivotTable pivotTable = sheet.PivotTables[0];
            IEnumerator pivotFieldEnum = pivotTable.RowFields.GetEnumerator();
            Console.WriteLine("\nPivot Table Row Fields:");
            while (pivotFieldEnum.MoveNext())
            {
                PivotField field = (PivotField)pivotFieldEnum.Current;
                Console.WriteLine($"Field Name: {field.Name}");
                IEnumerator itemEnum = field.PivotItems.GetEnumerator();
                Console.WriteLine("  Items:");
                while (itemEnum.MoveNext())
                {
                    PivotItem item = (PivotItem)itemEnum.Current;
                    Console.WriteLine($"    {item.Value}");
                }
            }

            // ------------------------------------------------------------
            // Scenario 5: Enumerate digital signatures (if any) for security auditing.
            // ------------------------------------------------------------
            // Note: Digital signature APIs are not available in the current Aspose.Cells version.
            // This section is omitted to ensure successful compilation.

            // Save the workbook after any modifications (if any were made)
            workbook.Save("ProcessedSampleData.xlsx");
        }
    }
}