using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsStrongTypedColumns
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ------------------------------------------------------------
            // Define column headers
            // ------------------------------------------------------------
            cells[0, 0].PutValue("ProductId");   // Integer column
            cells[0, 1].PutValue("ProductName"); // String column
            cells[0, 2].PutValue("Price");       // Double column
            cells[0, 3].PutValue("ReleaseDate"); // Date column

            // ------------------------------------------------------------
            // Apply data validation to enforce strong types per column
            // ------------------------------------------------------------
            // Integer validation for column A (ProductId)
            CellArea intArea = new CellArea { StartRow = 1, EndRow = 1000, StartColumn = 0, EndColumn = 0 };
            int intValidationIdx = sheet.Validations.Add(intArea);
            Validation intValidation = sheet.Validations[intValidationIdx];
            intValidation.Type = ValidationType.WholeNumber;
            intValidation.Operator = OperatorType.Between;
            intValidation.Formula1 = "1";
            intValidation.Formula2 = "1000000";
            intValidation.ShowError = true;
            intValidation.ErrorTitle = "Invalid ProductId";
            intValidation.ErrorMessage = "Enter an integer between 1 and 1,000,000.";

            // Text length validation for column B (ProductName)
            CellArea textArea = new CellArea { StartRow = 1, EndRow = 1000, StartColumn = 1, EndColumn = 1 };
            int textValidationIdx = sheet.Validations.Add(textArea);
            Validation textValidation = sheet.Validations[textValidationIdx];
            textValidation.Type = ValidationType.TextLength;
            textValidation.Operator = OperatorType.Between;
            textValidation.Formula1 = "1";
            textValidation.Formula2 = "50";
            textValidation.ShowError = true;
            textValidation.ErrorTitle = "Invalid ProductName";
            textValidation.ErrorMessage = "Enter a name between 1 and 50 characters.";

            // Decimal validation for column C (Price)
            CellArea doubleArea = new CellArea { StartRow = 1, EndRow = 1000, StartColumn = 2, EndColumn = 2 };
            int doubleValidationIdx = sheet.Validations.Add(doubleArea);
            Validation doubleValidation = sheet.Validations[doubleValidationIdx];
            doubleValidation.Type = ValidationType.Decimal;
            doubleValidation.Operator = OperatorType.Between;
            doubleValidation.Formula1 = "0";
            doubleValidation.Formula2 = "100000";
            doubleValidation.ShowError = true;
            doubleValidation.ErrorTitle = "Invalid Price";
            doubleValidation.ErrorMessage = "Enter a price between 0 and 100,000.";

            // Date validation for column D (ReleaseDate)
            CellArea dateArea = new CellArea { StartRow = 1, EndRow = 1000, StartColumn = 3, EndColumn = 3 };
            int dateValidationIdx = sheet.Validations.Add(dateArea);
            Validation dateValidation = sheet.Validations[dateValidationIdx];
            dateValidation.Type = ValidationType.Date;
            dateValidation.Operator = OperatorType.Between;
            dateValidation.Formula1 = "DATE(2000,1,1)";
            dateValidation.Formula2 = "DATE(2100,12,31)";
            dateValidation.ShowError = true;
            dateValidation.ErrorTitle = "Invalid ReleaseDate";
            dateValidation.ErrorMessage = "Enter a date between 01/01/2000 and 12/31/2100.";

            // ------------------------------------------------------------
            // Populate sample data using ImportObjectArray (vertical import)
            // ------------------------------------------------------------
            object[] sampleData = new object[]
            {
                // Row 1
                101, "Widget", 19.99, new DateTime(2023, 5, 1),
                // Row 2
                102, "Gadget", 29.49, new DateTime(2023, 6, 15),
                // Row 3
                103, "Doohickey", 9.95, new DateTime(2023, 7, 30)
            };
            // Import data starting at row 1 (index 1) because row 0 holds headers
            cells.ImportObjectArray(sampleData, 1, 0, false);

            // ------------------------------------------------------------
            // Export the worksheet to a DataTable with type checking
            // ------------------------------------------------------------
            ExportTableOptions exportOptions = new ExportTableOptions
            {
                ExportColumnName = true,          // Use first row as column names
                CheckMixedValueType = true,       // Ensure column types are inferred correctly
                PlotVisibleCells = true,
                PlotVisibleRows = true,
                PlotVisibleColumns = true
            };

            // Export rows 0-3 (header + 3 data rows) and 4 columns
            DataTable dt = sheet.Cells.ExportDataTable(0, 0, 4, 4, exportOptions);

            // Display exported column types (for verification)
            Console.WriteLine("Exported DataTable column types:");
            foreach (DataColumn col in dt.Columns)
            {
                Console.WriteLine($"{col.ColumnName}: {col.DataType}");
            }

            // ------------------------------------------------------------
            // Save the workbook (lifecycle rule: save)
            // ------------------------------------------------------------
            workbook.Save("StrongTypedColumnsDemo.xlsx");
        }
    }
}