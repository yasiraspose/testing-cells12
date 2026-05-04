using System;
using System.Data;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Build a sample DataTable that will serve as the data source
        DataTable dataTable = new DataTable();
        dataTable.Columns.Add("Product", typeof(string));
        dataTable.Columns.Add("Quantity", typeof(int));
        dataTable.Columns.Add("Price", typeof(decimal));

        dataTable.Rows.Add("Apple", 10, 0.5m);
        dataTable.Rows.Add("Banana", 20, 0.3m);
        dataTable.Rows.Add("Cherry", 15, 1.2m);

        // Write column headers
        for (int col = 0; col < dataTable.Columns.Count; col++)
        {
            cells[0, col].PutValue(dataTable.Columns[col].ColumnName);
        }

        // Write data rows
        for (int row = 0; row < dataTable.Rows.Count; row++)
        {
            for (int col = 0; col < dataTable.Columns.Count; col++)
            {
                cells[row + 1, col].PutValue(dataTable.Rows[row][col]);
            }
        }

        // Save the workbook as an XLSX file
        workbook.Save("DataGridImportDemo.xlsx");
    }
}