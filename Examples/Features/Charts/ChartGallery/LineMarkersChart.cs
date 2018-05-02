using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Charts.ChartGallery
{
    public class LineMarkersChart :ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            GrapeCity.Documents.Excel.Drawing.IShape shape = worksheet.Shapes.AddChart(GrapeCity.Documents.Excel.Drawing.ChartType.LineMarkers, 250, 20, 360, 230);
            worksheet.Range["A1:B8"].Value = new object[,] {
                { 6, 55 },
                { 45, 25 },
                { 35, 45 },
                { 25, 65 },
                { 65, 15 },
                { 45, 75 },
                { 75, 55 },
                { 65, 35 },
            };
            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:B8"], GrapeCity.Documents.Excel.Drawing.RowCol.Columns);
            shape.Chart.ChartTitle.Text = "Line with Markers";
            shape.Chart.SeriesCollection[0].MarkerStyle = GrapeCity.Documents.Excel.Drawing.MarkerStyle.Square;
        }
    }
}
