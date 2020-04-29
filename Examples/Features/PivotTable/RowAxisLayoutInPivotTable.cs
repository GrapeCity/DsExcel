using System;

namespace GrapeCity.Documents.Excel.Examples.Features.PivotTable
{
    public class RowAxisLayoutInPivotTable : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            object[,] sourceData = new object[,] {
                { "Order ID", "Product",  "Category",   "Amount", "Date",                    "Country" },
                { 1,          "Carrots",  "Vegetables",  4270,    new DateTime(2018, 1, 6),  "United States" },
                { 2,          "Broccoli", "Vegetables",  8239,    new DateTime(2018, 1, 7),  "United Kingdom" },
                { 3,          "Banana",   "Fruit",       617,     new DateTime(2018, 1, 8),  "United States" },
                { 4,          "Banana",   "Fruit",       8384,    new DateTime(2018, 1, 10), "Canada" },
                { 5,          "Beans",    "Vegetables",  2626,    new DateTime(2018, 1, 10), "Germany" },
                { 6,          "Orange",   "Fruit",       3610,    new DateTime(2018, 1, 11), "United States" },
                { 7,          "Broccoli", "Vegetables",  9062,    new DateTime(2018, 1, 11), "Australia" },
                { 8,          "Banana",   "Fruit",       6906,    new DateTime(2018, 1, 16), "New Zealand" },
                { 9,          "Apple",    "Fruit",       2417,    new DateTime(2018, 1, 16), "France" },
                { 10,         "Apple",    "Fruit",       7431,    new DateTime(2018, 1, 16), "Canada" },
                { 11,         "Banana",   "Fruit",       8250,    new DateTime(2018, 1, 16), "Germany" },
                { 12,         "Broccoli", "Vegetables",  7012,    new DateTime(2018, 1, 18), "United States" },
                { 13,         "Carrots",  "Vegetables",  1903,    new DateTime(2018, 1, 20), "Germany" },
                { 14,         "Broccoli", "Vegetables",  2824,    new DateTime(2018, 1, 22), "Canada" },
                { 15,         "Apple",    "Fruit",       6946,    new DateTime(2018, 1, 24), "France" },
            };

            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Range["A1:F16"].Value = sourceData;
            worksheet.Range["A:F"].ColumnWidth = 15;
            var pivotcache = workbook.PivotCaches.Create(worksheet.Range["A1:F16"]);
            var pivottable = worksheet.PivotTables.Add(pivotcache, worksheet.Range["H7"], "pivottable1");
            worksheet.Range["D2:D16"].NumberFormat = "$#,##0.00";
            worksheet.Range["I9:O11"].NumberFormat = "$#,##0.00";
            worksheet.Range["H:O"].ColumnWidth = 12;

            //config pivot table's fields
            var field_Category = pivottable.PivotFields["Category"];
            field_Category.Orientation = PivotFieldOrientation.RowField;

            var field_Product = pivottable.PivotFields["Product"];
            field_Product.Orientation = PivotFieldOrientation.ColumnField;

            var field_Amount = pivottable.PivotFields["Amount"];
            field_Amount.Orientation = PivotFieldOrientation.DataField;

            var field_Country = pivottable.PivotFields["Country"];
            field_Country.Orientation = PivotFieldOrientation.RowField;

            // Set row axis layout to tabular
            pivottable.SetRowAxisLayout(LayoutRowType.TabularRow);
        }

        public override bool ShowViewer
        {
            get
            {
                return false;
            }
        }

        public override bool ShowScreenshot
        {
            get
            {
                return true;
            }
        }
    }
}
