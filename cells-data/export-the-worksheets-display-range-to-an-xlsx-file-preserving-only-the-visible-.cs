using System;
using System.Data;
using Aspose.Cells;

namespace ExportVisibleCellsDemo
{
    class Program
    {
        static void Main()
        {
            // -------------------------------------------------
            // 1. Create a source workbook and populate sample data
            // -------------------------------------------------
            Workbook sourceWb = new Workbook();
            Worksheet srcSheet = sourceWb.Worksheets[0];
            Cells srcCells = srcSheet.Cells;

            // Fill a 5x5 range with sample values
            for (int row = 0; row < 5; row++)
                for (int col = 0; col < 5; col++)
                    srcCells[row, col].PutValue($"R{row + 1}C{col + 1}");

            // Hide some rows and columns to simulate non‑visible cells
            srcSheet.Cells.HideRow(1);      // hide second row
            srcSheet.Cells.HideColumn(2);   // hide third column

            // -------------------------------------------------
            // 2. Export only the visible cells to a DataTable
            // -------------------------------------------------
            ExportTableOptions exportOpts = new ExportTableOptions
            {
                PlotVisibleCells = true,      // export only visible cells
                PlotVisibleRows = true,
                PlotVisibleColumns = true,
                ExportColumnName = true       // include column headers
            };

            // Export the whole used range (0,0) with 5 rows and 5 columns
            DataTable visibleData = srcSheet.Cells.ExportDataTable(0, 0, 5, 5, exportOpts);

            // -------------------------------------------------
            // 3. Create a new workbook and import the DataTable
            // -------------------------------------------------
            Workbook targetWb = new Workbook();
            Worksheet targetSheet = targetWb.Worksheets[0];
            Cells targetCells = targetSheet.Cells;

            // Manually write the DataTable to the target sheet (including headers)
            int startRow = 0;
            int startCol = 0;

            // Write column headers
            for (int c = 0; c < visibleData.Columns.Count; c++)
            {
                targetCells[startRow, startCol + c].PutValue(visibleData.Columns[c].ColumnName);
            }

            // Write data rows
            for (int r = 0; r < visibleData.Rows.Count; r++)
            {
                for (int c = 0; c < visibleData.Columns.Count; c++)
                {
                    targetCells[startRow + 1 + r, startCol + c].PutValue(visibleData.Rows[r][c]);
                }
            }

            // -------------------------------------------------
            // 4. Save the result as XLSX – only visible cells are present
            // -------------------------------------------------
            targetWb.Save("VisibleCellsExport.xlsx", SaveFormat.Xlsx);
        }
    }
}