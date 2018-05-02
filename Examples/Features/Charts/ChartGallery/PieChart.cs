using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Charts.ChartGallery
{
    public class PieChart : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            GrapeCity.Documents.Excel.Drawing.IShape shape = worksheet.Shapes.AddChart(GrapeCity.Documents.Excel.Drawing.ChartType.Pie, 250, 20, 360, 230);
            worksheet.Range["A1:B4"].Value = new object[,] {
                { "Blue", 1 },
                { "Red", 2 },
                { "Green", 3 },
                { "Purple", 4 },             
            };
            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:B4"], GrapeCity.Documents.Excel.Drawing.RowCol.Columns);
            shape.Chart.ChartTitle.Text = "Pie Chart";
            shape.Chart.Legend.Position = GrapeCity.Documents.Excel.Drawing.LegendPosition.Right;
        }
    }
}
