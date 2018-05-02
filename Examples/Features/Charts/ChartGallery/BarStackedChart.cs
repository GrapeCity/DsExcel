using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Charts.ChartGallery
{
    public class BarStackedChart :ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            GrapeCity.Documents.Excel.Drawing.IShape shape = worksheet.Shapes.AddChart(GrapeCity.Documents.Excel.Drawing.ChartType.BarStacked, 250, 20, 360, 230);
            worksheet.Range["A1:C4"].Value = new object[,] {
                { 103, 121, 109 },
                { 56, 94, 115 },
                { 116, 89, 99 },
                { 55, 93, 70 }             
            };
            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:C4"], GrapeCity.Documents.Excel.Drawing.RowCol.Columns);
            shape.Chart.ChartTitle.Text = "Bar Stacked Chart";
            shape.Chart.Legend.Position = GrapeCity.Documents.Excel.Drawing.LegendPosition.Left;
        }
    }
}
