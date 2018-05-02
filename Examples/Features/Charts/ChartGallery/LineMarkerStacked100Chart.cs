using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Charts.ChartGallery
{
    public class LineMarkerStacked100Chart :ExampleBase 
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            GrapeCity.Documents.Excel.Drawing.IShape shape = worksheet.Shapes.AddChart(GrapeCity.Documents.Excel.Drawing.ChartType.LineMarkersStacked100, 250, 20, 360, 230);
            worksheet.Range["A1:C5"].Value = new object[,]
            {
                {12, 22, 27},
                {45, 52, 25},
                {58, 35, 58},
                {21, 37, 43},
                {44, 45, 28}
            };
            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:C5"], GrapeCity.Documents.Excel.Drawing.RowCol.Columns);
            shape.Chart.ChartTitle.Text = "Line Marker Stacked 100 Chart";
        }
    }
}
