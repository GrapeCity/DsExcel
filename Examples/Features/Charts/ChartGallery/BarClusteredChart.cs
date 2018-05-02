using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Charts.ChartGallery
{
   public class BarClusteredChart : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            GrapeCity.Documents.Excel.Drawing.IShape shape = worksheet.Shapes.AddChart(GrapeCity.Documents.Excel.Drawing.ChartType.BarClustered, 250, 20, 360, 230);
            worksheet.Range["A1:D2"].Value = new object[,] {
                { 100,200,300,400 },
                { 100,200,300,400 },            
            };
            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:D2"], GrapeCity.Documents.Excel.Drawing.RowCol.Columns);
            shape.Chart.ChartTitle.Text = "Bar Clustered Chart";
            shape.Chart.Legend.Position = GrapeCity.Documents.Excel.Drawing.LegendPosition.Left;
        }
    }
}
