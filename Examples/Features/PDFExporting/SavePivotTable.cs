using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting
{
    public class SavePivotTable : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            object[,] sourceData = new object[,] {
                { "Order ID", "Product",  "Category",   "Amount", "Date",                    "Country" },
                { 1,          "Broccoli", "Vegetables",  8239,    new DateTime(2018, 1, 7),  "United Kingdom" },
                { 2,          "Banana",   "Fruit",       617,     new DateTime(2018, 1, 8),  "United States" },
                { 3,          "Banana",   "Fruit",       8384,    new DateTime(2018, 1, 10), "Canada" },
                { 4,          "Beans",    "Vegetables",  2626,    new DateTime(2018, 1, 10), "Germany" },
                { 5,          "Orange",   "Fruit",       3610,    new DateTime(2018, 1, 11), "United States" },
                { 6,          "Broccoli", "Vegetables",  9062,    new DateTime(2018, 1, 11), "Australia" },
                { 7,          "Banana",   "Fruit",       6906,    new DateTime(2018, 1, 16), "New Zealand" },
                { 8,          "Apple",    "Fruit",       2417,    new DateTime(2018, 1, 16), "France" },
                { 9,         "Apple",    "Fruit",       7431,    new DateTime(2018, 1, 16), "Canada" },
                { 10,         "Banana",   "Fruit",       8250,    new DateTime(2018, 1, 16), "Germany" },
                { 11,         "Broccoli", "Vegetables",  7012,    new DateTime(2018, 1, 18), "United States" },
                { 12,         "Broccoli", "Vegetables",  2824,    new DateTime(2018, 1, 22), "Canada" },
                { 13,         "Apple",    "Fruit",       6946,    new DateTime(2018, 1, 24), "France" },
            };

            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Range["K20:P33"].Value = sourceData;
            worksheet.Range["K:P"].ColumnWidth = 15;
            // Add pivot table
            var pivotcache = workbook.PivotCaches.Create(worksheet.Range["K20:P33"]);
            var pivottable = worksheet.PivotTables.Add(pivotcache, worksheet.Range["A1"], "pivottable1");
            worksheet.Range["N21:N35"].NumberFormat = "$#,##0.00";
            worksheet.Range["A:G"].ColumnWidth = 12;

            //config pivot table's fields
            var field_Date = pivottable.PivotFields["Date"];
            field_Date.Orientation = PivotFieldOrientation.PageField;

            var field_Category = pivottable.PivotFields["Category"];
            field_Category.Orientation = PivotFieldOrientation.RowField;

            var field_Product = pivottable.PivotFields["Product"];
            field_Product.Orientation = PivotFieldOrientation.ColumnField;

            var field_Amount = pivottable.PivotFields["Amount"];
            field_Amount.Orientation = PivotFieldOrientation.DataField;
            field_Amount.NumberFormat = "$#,##0.00";

            var field_Country = pivottable.PivotFields["Country"];
            field_Country.Orientation = PivotFieldOrientation.RowField;

            // Set pivot style
            pivottable.TableStyle = "PivotStyleMedium28";
        }

        public override bool SavePdf
        {
            get
            {
                return true;
            }
        }
    }
}
