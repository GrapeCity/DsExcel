using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Charts.ChartGallery
{
    public class XYScatterChart :ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            GrapeCity.Documents.Excel.Drawing.IShape shape = worksheet.Shapes.AddChart(GrapeCity.Documents.Excel.Drawing.ChartType.XYScatter, 250, 20, 360, 230);
            worksheet.Range["A1:D7"].Value = new object[,] {
                { "Blue", null, "Red", null },
                { 55, 964, 67, 475 },
                { 20, 825, 10, 163 },
                { 77, 840, 87, 224 },
                { 182, 596, 46, 196 },
                { 190, 384, 100, 377 },
                { 140, 503, 92, 47 },
            };
            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:B7"], GrapeCity.Documents.Excel.Drawing.RowCol.Columns);
            shape.Chart.SeriesCollection.Add(worksheet.Range["C1:D7"], GrapeCity.Documents.Excel.Drawing.RowCol.Columns);
            shape.Chart.ChartTitle.Text = "Scatter Chart";
            //config markers style
            GrapeCity.Documents.Excel.Drawing.ISeries series1 = shape.Chart.SeriesCollection[0];
            GrapeCity.Documents.Excel.Drawing.ISeries series2 = shape.Chart.SeriesCollection[1];
            series1.MarkerStyle = Drawing.MarkerStyle.Square;
            series1.MarkerSize = 10;
            series2.MarkerSize = 10;
        }
    }
}
