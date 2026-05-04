using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Settings;

namespace AsposeCellsExamples
{
    public class PivotLocalizationDemo
    {
        public static void Run()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a settable globalization settings instance
            SettableGlobalizationSettings globalizationSettings = new SettableGlobalizationSettings();

            // Create a settable pivot globalization settings instance
            SettablePivotGlobalizationSettings pivotSettings = new SettablePivotGlobalizationSettings();

            // Choose the language for localization (e.g., "en", "fr", "de")
            string language = "fr";

            // Apply language‑specific texts for Total, Grand Total and Subtotals
            if (language == "fr")
            {
                // Total label
                pivotSettings.SetTextOfTotal("Total");

                // Grand Total label
                pivotSettings.SetTextOfGrandTotal("Total général");

                // Subtotal labels
                pivotSettings.SetTextOfSubTotal(PivotFieldSubtotalType.Sum, "Somme");
                pivotSettings.SetTextOfSubTotal(PivotFieldSubtotalType.Count, "Nombre");
                pivotSettings.SetTextOfSubTotal(PivotFieldSubtotalType.Average, "Moyenne");
                pivotSettings.SetTextOfSubTotal(PivotFieldSubtotalType.Max, "Maximum");
                pivotSettings.SetTextOfSubTotal(PivotFieldSubtotalType.Min, "Minimum");
            }
            else if (language == "de")
            {
                // Total label
                pivotSettings.SetTextOfTotal("Summe");

                // Grand Total label
                pivotSettings.SetTextOfGrandTotal("Gesamtsumme");

                // Subtotal labels
                pivotSettings.SetTextOfSubTotal(PivotFieldSubtotalType.Sum, "Summe");
                pivotSettings.SetTextOfSubTotal(PivotFieldSubtotalType.Count, "Anzahl");
                pivotSettings.SetTextOfSubTotal(PivotFieldSubtotalType.Average, "Durchschnitt");
                pivotSettings.SetTextOfSubTotal(PivotFieldSubtotalType.Max, "Maximum");
                pivotSettings.SetTextOfSubTotal(PivotFieldSubtotalType.Min, "Minimum");
            }
            // else default English – no changes required

            // Attach the pivot globalization settings to the workbook's globalization settings
            globalizationSettings.PivotSettings = pivotSettings;
            workbook.Settings.GlobalizationSettings = globalizationSettings;

            // (Optional) Create a pivot table to demonstrate the effect
            // Assuming the worksheet contains data in A1:B5
            int pivotIndex = worksheet.PivotTables.Add("A1:B5", "D1", "LocalizedPivot");
            PivotTable pivotTable = worksheet.PivotTables[pivotIndex];
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            PivotLocalizationDemo.Run();
        }
    }
}