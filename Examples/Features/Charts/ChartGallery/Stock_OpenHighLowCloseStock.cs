using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Charts.ChartGallery
{
    public class Stock_OpenHighLowCloseStock : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            GrapeCity.Documents.Excel.Drawing.IShape shape = worksheet.Shapes.AddChart(GrapeCity.Documents.Excel.Drawing.ChartType.StockOHLC, 350, 20, 360, 220);
            worksheet.Range["A1:E17"].Value = new object[,] {
                    { null, "Open", "High", "Low", "Close" },
                    { new DateTime(2008, 9, 1), 103.46, 105.76, 92.38, 100.94 },
                    { new DateTime(2008, 9, 2), 100.26, 102.45, 90.14, 93.45 },
                    { new DateTime(2008, 9, 3), 98.05, 102.11, 85.01, 99.89 },
                    { new DateTime(2008, 9, 4), 100.32, 106.01, 94.04, 99.45 },
                    { new DateTime(2008, 9, 5), 99.74, 108.23, 98.16, 104.33 },
                    { new DateTime(2008, 9, 8), 92.11, 107.7, 91.02, 102.17 },
                    { new DateTime(2008, 9, 9), 107.8, 110.36, 101.62, 110.07 },
                    { new DateTime(2008, 9, 10), 107.56, 115.97, 106.89, 112.39 },
                    { new DateTime(2008, 9, 11), 112.86, 120.32, 112.15, 117.52 },
                    { new DateTime(2008, 9, 12), 115.02, 122.03, 114.67, 114.75 },
                    { new DateTime(2008, 9, 15), 108.53, 120.46, 106.21, 116.85 },
                    { new DateTime(2008, 9, 16), 114.97, 118.08, 113.55, 116.69 },
                    { new DateTime(2008, 9, 17), 127.14, 128.23, 110.91, 117.25 },
                    { new DateTime(2008, 9, 18), 118.89, 120.55, 108.09, 112.52 },
                    { new DateTime(2008, 9, 19), 105.57, 112.58, 105.42, 109.12 },
                    { new DateTime(2008, 9, 22), 110.23, 115.23, 97.25, 101.56 },
                };
            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:E17"], GrapeCity.Documents.Excel.Drawing.RowCol.Columns);
            //set chart title
            shape.Chart.ChartTitle.Text = "Open-High-Low-Close Stock Chart";
          
            GrapeCity.Documents.Excel.Drawing.IAxis valueAxis = shape.Chart.Axes.Item(GrapeCity.Documents.Excel.Drawing.AxisType.Value);
            GrapeCity.Documents.Excel.Drawing.IAxis categoryAxis = shape.Chart.Axes.Item(GrapeCity.Documents.Excel.Drawing.AxisType.Category);           
            //config value axis 
            valueAxis.MinimumScale = 80;
            valueAxis.MaximumScale = 140;
            valueAxis.MajorUnit = 15;
            //config category axis
            categoryAxis.CategoryType = Drawing.CategoryType.CategoryScale;
            categoryAxis.MajorTickMark = Drawing.TickMark.Outside;
            categoryAxis.TickMarkSpacing = 5;
            categoryAxis.TickLabelSpacing = 5;
        }
    }
}
