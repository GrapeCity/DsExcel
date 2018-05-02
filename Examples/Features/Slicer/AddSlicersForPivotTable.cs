using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Slicer
{
    public class AddSlicersForPivotTable : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            object[,] sourceData = new object[,] {
                { "Order ID", "Product",  "Category",   "Amount", "Date",                    "Country" },
                { 1,          "Carrots",  "Vegetables",  4270,    new DateTime(2012, 1, 6),  "United States" },
                { 2,          "Broccoli", "Vegetables",  8239,    new DateTime(2012, 1, 7),  "United Kingdom" },
                { 3,          "Banana",   "Fruit",       617,     new DateTime(2012, 1, 8),  "United States" },
                { 4,          "Banana",   "Fruit",       8384,    new DateTime(2012, 1, 10), "Canada" },
                { 5,          "Beans",    "Vegetables",  2626,    new DateTime(2012, 1, 10), "Germany" },
                { 6,          "Orange",   "Fruit",       3610,    new DateTime(2012, 1, 11), "United States" },
                { 7,          "Broccoli", "Vegetables",  9062,    new DateTime(2012, 1, 11), "Australia" },
                { 8,          "Banana",   "Fruit",       6906,    new DateTime(2012, 1, 16), "New Zealand" },
                { 9,          "Apple",    "Fruit",       2417,    new DateTime(2012, 1, 16), "France" },
                { 10,         "Apple",    "Fruit",       7431,    new DateTime(2012, 1, 16), "Canada" },
                { 11,         "Banana",   "Fruit",       8250,    new DateTime(2012, 1, 16), "Germany" },
                { 12,         "Broccoli", "Vegetables",  7012,    new DateTime(2012, 1, 18), "United States" },
                { 13,         "Carrots",  "Vegetables",  1903,    new DateTime(2012, 1, 20), "Germany" },
                { 14,         "Broccoli", "Vegetables",  2824,    new DateTime(2012, 1, 22), "Canada" },
                { 15,         "Apple",    "Fruit",       6946,    new DateTime(2012, 1, 24), "France" },
            };

            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Range["A1:F16"].Value = sourceData;
            worksheet.Range["A:F"].ColumnWidth = 15;

            //Create pivot cache.
            IPivotCache pivotcache = workbook.PivotCaches.Create(worksheet.Range["A1:F16"]);
            //Create pivot tables.
            IPivotTable pivottable1 = worksheet.PivotTables.Add(pivotcache, worksheet.Range["K5"], "pivottable1");
            IPivotTable pivottable2 = worksheet.PivotTables.Add(pivotcache, worksheet.Range["N3"], "pivottable2");

            //Config pivot fields
            IPivotField field_product1 = pivottable1.PivotFields[1];
            field_product1.Orientation = PivotFieldOrientation.RowField;

            IPivotField field_Amount1 = pivottable1.PivotFields[3];
            field_Amount1.Orientation = PivotFieldOrientation.DataField;

            IPivotField field_product2 = pivottable2.PivotFields[5];
            field_product2.Orientation = PivotFieldOrientation.RowField;

            IPivotField field_Amount2 = pivottable2.PivotFields[2];
            field_Amount2.Orientation = PivotFieldOrientation.DataField;
            field_Amount2.Function = ConsolidationFunction.Count;

            //create slicer cache, the slicers base the slicer cache just control pivot table1.
            ISlicerCache cache = workbook.SlicerCaches.Add(pivottable1, "Product");
            ISlicer slicer1 = cache.Slicers.Add(workbook.Worksheets["Sheet1"], "p1", "Product", 30, 550, 100, 200);

            //add pivot table2 for slicer cache, the slicers base the slicer cache will control pivot tabl1 and pivot table2.
            cache.PivotTables.AddPivotTable(pivottable2);
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
