using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Charts.ChartGallery
{
    public class XYScatterSmoothWithMarkers : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            GrapeCity.Documents.Excel.Drawing.IShape shape = worksheet.Shapes.AddChart(GrapeCity.Documents.Excel.Drawing.ChartType.XYScatterSmooth, 250, 20, 360, 230);
            worksheet.Range["A1:B5"].Value = new object[,] {
                { 4, 2 },
                { 6, 1 },
                { 1, 2 },
                { 7, 4 },
                { 4, 4 },
            };
            worksheet.Range["A7:B11"].Value = new object[,] {
                { 9, 5 },
                { 7, 8 },
                { 9, 8 },
                { 5, 9 },
                { 2, 4 },
            };
            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:B5"], GrapeCity.Documents.Excel.Drawing.RowCol.Columns);
            shape.Chart.SeriesCollection.Add(worksheet.Range["A7:B11"], GrapeCity.Documents.Excel.Drawing.RowCol.Columns);
            shape.Chart.ChartTitle.Text = "Scatter with Smooth Lines and Markers";
        }
    }
}
