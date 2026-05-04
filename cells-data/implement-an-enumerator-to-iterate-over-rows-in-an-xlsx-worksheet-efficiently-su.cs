using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells;

namespace RowEnumeratorDemo
{
    class Program
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // Example 1: Lazy row enumeration using RowCollection.GetEnumerator
            // ------------------------------------------------------------
            // Load a workbook (replace with the actual file path)
            LoadOptions loadOptions = new LoadOptions();
            Workbook workbook = new Workbook("LargeFile.xlsx", loadOptions);
            Worksheet worksheet = workbook.Worksheets[0];

            // Enumerate rows lazily; rows are instantiated only when accessed
            foreach (Row row in EnumerateRows(worksheet))
            {
                // Access only the data you need to keep memory usage low
                Cell firstCell = row.GetCellOrNull(0);
                string value = firstCell != null ? firstCell.StringValue : "<empty>";
                Console.WriteLine($"Row {row.Index}: {value}");
            }

            // ------------------------------------------------------------
            // Example 2: Streaming processing with LightCellsDataHandler
            // ------------------------------------------------------------
            // The handler processes rows one by one while the workbook is being loaded,
            // avoiding the need to keep the whole worksheet in memory.
            LoadOptions streamingOptions = new LoadOptions();
            streamingOptions.LightCellsDataHandler = new StreamingRowHandler();

            // Loading the workbook triggers the handler callbacks
            Workbook streamingWorkbook = new Workbook("LargeFile.xlsx", streamingOptions);
            // No further code required; the handler does the work during load
        }

        // Lazy enumerator that wraps the built‑in RowCollection enumerator.
        // reversed = false, sync = false gives the best performance.
        private static IEnumerable<Row> EnumerateRows(Worksheet sheet, bool reversed = false, bool sync = false)
        {
            IEnumerator enumerator = sheet.Cells.Rows.GetEnumerator(reversed, sync);
            while (enumerator.MoveNext())
            {
                yield return (Row)enumerator.Current;
            }
        }

        // LightCellsDataHandler implementation for true streaming processing.
        private class StreamingRowHandler : LightCellsDataHandler
        {
            public bool StartSheet(Worksheet sheet)
            {
                Console.WriteLine($"Processing sheet: {sheet.Name}");
                return true; // Process all rows in this sheet
            }

            public bool StartRow(int rowIndex)
            {
                // Return true to indicate that this row (and its cells) should be processed
                Console.WriteLine($"Start processing row {rowIndex}");
                return true;
            }

            public bool ProcessRow(Row row)
            {
                // Row object is provided; read only the first cell to keep memory low
                Cell cell = row.GetCellOrNull(0);
                string val = cell != null ? cell.StringValue : "<empty>";
                Console.WriteLine($"Row {row.Index} first cell: {val}");
                return true; // Continue processing cells in this row
            }

            public bool StartCell(int columnIndex)
            {
                // Process all cells; return true to receive ProcessCell callbacks
                return true;
            }

            public bool ProcessCell(Cell cell)
            {
                // Example: output cell address and value
                Console.WriteLine($"Cell {cell.Name}: {cell.StringValue}");
                return true;
            }
        }
    }
}