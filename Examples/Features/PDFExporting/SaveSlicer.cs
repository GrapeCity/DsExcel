using System;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting
{
    public class SaveSlicer : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            object[,] sourceData = new object[,] {
                { "Order ID", "Product",  "Category",   "Amount", "Date"                    },
                { 1,          "Carrots",  "Vegetables",  4270,    new DateTime(2018, 1, 6)  },
                { 2,          "Broccoli", "Vegetables",  8239,    new DateTime(2018, 1, 7)  },
                { 3,          "Banana",   "Fruit",       617,     new DateTime(2018, 1, 8)  },
                { 4,          "Banana",   "Fruit",       8384,    new DateTime(2018, 1, 10) },
                { 5,          "Beans",    "Vegetables",  2626,    new DateTime(2018, 1, 10) },
                { 6,          "Orange",   "Fruit",       3610,    new DateTime(2018, 1, 11) },
                { 7,          "Broccoli", "Vegetables",  9062,    new DateTime(2018, 1, 11) },
                { 8,          "Banana",   "Fruit",       6906,    new DateTime(2018, 1, 16) },
                { 9,          "Apple",    "Fruit",       2417,    new DateTime(2018, 1, 16) },
                { 10,         "Apple",    "Fruit",       7431,    new DateTime(2018, 1, 16) },
                { 11,         "Banana",   "Fruit",       8250,    new DateTime(2018, 1, 16) },
                { 12,         "Broccoli", "Vegetables",  7012,    new DateTime(2018, 1, 18) },
                { 13,         "Carrots",  "Vegetables",  1903,    new DateTime(2018, 1, 20) },
                { 14,         "Broccoli", "Vegetables",  2824,    new DateTime(2018, 1, 22) },
                { 15,         "Apple",    "Fruit",       6946,    new DateTime(2018, 1, 24) },
            };

            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Range["A:E"].ColumnWidth = 15;

            worksheet.Range["A1:E16"].Value = sourceData;
            ITable table = worksheet.Tables.Add(worksheet.Range["A1:E16"], true);
            table.Columns[3].DataBodyRange.NumberFormat = "$#,##0.00";
            //create slicer cache for table.
            ISlicerCache cache = workbook.SlicerCaches.Add(table, "Category", "categoryCache");

            //add slicer for Category column.
            ISlicer slicer1 = cache.Slicers.Add(workbook.Worksheets["Sheet1"], "cate1", "Category", 150, 30, 100, 200);
            slicer1.SlicerCache.SlicerItems["Vegetables"].Selected = false;
        }

        public override bool SavePdf
        {
            get
            {
                return true;
            }
        }

        public override bool ShowViewer
        {
            get
            {
                return false;
            }
        }

        public override bool IsNew
        {
            get
            {
                return true;
            }
        }
    }
}
