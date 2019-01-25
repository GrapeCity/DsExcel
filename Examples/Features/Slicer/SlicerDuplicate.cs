using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Slicer
{
    public class SlicerDuplicate : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

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

            worksheet.Range["A:F"].ColumnWidth = 15;

            worksheet.Range["A1:F16"].Value = sourceData;
            ITable table = worksheet.Tables.Add(worksheet.Range["A1:F16"], true);
            table.Columns[3].DataBodyRange.NumberFormat = "$#,##0.00";

            //Create slicer cache for table.
            ISlicerCache cache = workbook.SlicerCaches.Add(table, "Category", "categoryCache");

            //Add slicer
            ISlicer slicer = cache.Slicers.Add(workbook.Worksheets["Sheet1"], "cate1", "Category", 30, 550, 100, 200);

            //Duplicate slicer
            GrapeCity.Documents.Excel.Drawing.IShape newShape = slicer.Shape.Duplicate();

        }
        
    }
}
