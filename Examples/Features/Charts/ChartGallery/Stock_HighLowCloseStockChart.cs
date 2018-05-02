using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Charts.ChartGallery
{
    public class Stock_HighLowCloseStockChart :ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            GrapeCity.Documents.Excel.Drawing.IShape shape = worksheet.Shapes.AddChart(GrapeCity.Documents.Excel.Drawing.ChartType.StockHLC, 350, 20, 360, 230);
            worksheet.Range["A1:D17"].Value = new object[,] {
                { null, "High", "Low", "Close" },
                { new DateTime(2019, 9, 1), 105.76, 92.38, 100.94 },
                { new DateTime(2019, 9, 2), 102.45, 90.14, 93.45 },
                { new DateTime(2019, 9, 3),102.11, 85.01, 99.89 },
                { new DateTime(2019, 9, 4), 106.01, 94.04, 99.45 },
                { new DateTime(2019, 9, 5),108.23, 98.16, 104.33 },
                { new DateTime(2019, 9, 8),107.7, 91.02, 102.17 },
                { new DateTime(2019, 9, 9),110.36, 101.62, 110.07 },
                { new DateTime(2019, 9, 10),115.97, 106.89, 112.39 },
                { new DateTime(2019, 9, 11),120.32, 112.15, 117.52 },
                { new DateTime(2019, 9, 12),122.03, 114.67, 114.75 },
                { new DateTime(2019, 9, 15),120.46, 106.21, 116.85 },
                { new DateTime(2019, 9, 16),118.08, 113.55, 116.69 },
                { new DateTime(2019, 9, 17),128.23, 110.91, 117.25 },
                { new DateTime(2019, 9, 18),120.55, 108.09, 112.52 },
                { new DateTime(2019, 9, 19),112.58, 105.42, 109.12 },
                { new DateTime(2019, 9, 22),115.23, 97.25, 101.56 },
            };
            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:D17"], GrapeCity.Documents.Excel.Drawing.RowCol.Columns);
            shape.Chart.ChartTitle.Text = "High-Low-Close Stock Chart";
            GrapeCity.Documents.Excel.Drawing.IAxis valueAxis = shape.Chart.Axes.Item(GrapeCity.Documents.Excel.Drawing.AxisType.Value);
            GrapeCity.Documents.Excel.Drawing.IAxis categoryAxis = shape.Chart.Axes.Item(GrapeCity.Documents.Excel.Drawing.AxisType.Category);
            GrapeCity.Documents.Excel.Drawing.ISeries series_close = shape.Chart.SeriesCollection[2];
            //config value axis
            valueAxis.MinimumScale = 80;
            valueAxis.MaximumScale = 140;
            valueAxis.MajorUnit = 15;
            //config category axis
            categoryAxis.CategoryType = Drawing.CategoryType.CategoryScale;
            categoryAxis.MajorTickMark = Drawing.TickMark.Outside;
            categoryAxis.TickLabelSpacingIsAuto = false;
            categoryAxis.TickLabelSpacing = 5;
            series_close.MarkerStyle = Drawing.MarkerStyle.Square;
        }
    }
}
