using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Charts.ChartGallery
{
    public class Stock_VolumeOpenHighLowClose:ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            GrapeCity.Documents.Excel.Drawing.IShape shape = worksheet.Shapes.AddChart(GrapeCity.Documents.Excel.Drawing.ChartType.StockVOHLC, 300, 20, 360, 230);
            worksheet.Range["A1:F23"].Value = new object[,] {
                { null, "Volume", "Open", "High", "Low", "Close" },
                { new DateTime(2019, 9, 1), 26085, 103.46, 105.76, 92.38, 100.94 },
                { new DateTime(2019, 9, 2), 52314, 100.26, 102.45, 90.14, 93.45 },
                { new DateTime(2019, 9, 3), 70308, 98.05, 102.11, 85.01, 99.89 },
                { new DateTime(2019, 9, 4), 33401, 100.32, 106.01, 94.04, 99.45 },
                { new DateTime(2019, 9, 5), 87500, 99.74, 108.23, 98.16, 104.33 },
                { new DateTime(2019, 9, 8), 33756, 92.11, 107.7, 91.02, 102.17 },
                { new DateTime(2019, 9, 9), 65737, 107.8, 110.36, 101.62, 110.07 },
                { new DateTime(2019, 9, 10), 45668, 107.56, 115.97, 106.89, 112.39 },
                { new DateTime(2019, 9, 11), 47815, 112.86, 120.32, 112.15, 117.52 },
                { new DateTime(2019, 9, 12), 76759, 115.02, 122.03, 114.67, 114.75 },
                { new DateTime(2019, 9, 15), 23492, 108.53, 120.46, 106.21, 116.85 },
                { new DateTime(2019, 9, 16), 56127, 114.97, 118.08, 113.55, 116.69 },
                { new DateTime(2019, 9, 17), 81142, 127.14, 128.23, 110.91, 117.25 },
                { new DateTime(2019, 9, 18), 46384, 118.89, 120.55, 108.09, 112.52 },
                { new DateTime(2019, 9, 19), 51005, 105.57, 112.58, 105.42, 109.12 },
            };
            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:F23"], GrapeCity.Documents.Excel.Drawing.RowCol.Columns);
            shape.Chart.ChartTitle.Text = "Stock Volume-Open-High-Low-Close Chart";
            GrapeCity.Documents.Excel.Drawing.IAxis valueAxis = shape.Chart.Axes.Item(GrapeCity.Documents.Excel.Drawing.AxisType.Value);
            GrapeCity.Documents.Excel.Drawing.IAxis categoryAxis = shape.Chart.Axes.Item(GrapeCity.Documents.Excel.Drawing.AxisType.Category);
            GrapeCity.Documents.Excel.Drawing.IAxis valueSecondaryAxis = shape.Chart.Axes.Item(GrapeCity.Documents.Excel.Drawing.AxisType.Value, GrapeCity.Documents.Excel.Drawing.AxisGroup.Secondary);
            valueAxis.MinimumScale = 0;
            valueAxis.MaximumScale = 150000;
            valueAxis.MajorUnit = 30000;
            categoryAxis.CategoryType = GrapeCity.Documents.Excel.Drawing.CategoryType.CategoryScale;
            categoryAxis.TickLabelSpacing = 5;
            valueSecondaryAxis.MajorUnit = 40;
        }
    }
}
