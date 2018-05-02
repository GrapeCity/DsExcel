using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Charts.ChartGallery
{
    public class BubbleChart : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            GrapeCity.Documents.Excel.Drawing.IShape shape = worksheet.Shapes.AddChart(GrapeCity.Documents.Excel.Drawing.ChartType.Bubble, 250, 20, 360, 230);
            worksheet.Range["A1:C10"].Value = new object[,] {
                {"Blue", null, null },
                {125, 750, 3 },
                {25, 625, 7 },
                {75, 875, 5 },
                {175, 625, 6},
                {"Red",null,null },
                {125 ,500 , 10 },
                {25, 250, 1 },
                {75, 125, 5 },
                {175, 250, 8 },
            };
            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:C5"], GrapeCity.Documents.Excel.Drawing.RowCol.Columns);
            shape.Chart.SeriesCollection.Add(worksheet.Range["A6:C10"], GrapeCity.Documents.Excel.Drawing.RowCol.Columns);
            shape.Chart.ChartTitle.Text = "Bubble Chart";
        }
    }
}
