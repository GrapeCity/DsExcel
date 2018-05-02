using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Charts.ChartGallery
{
    public class ColumnStacked100Chart :ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            GrapeCity.Documents.Excel.Drawing.IShape shape = worksheet.Shapes.AddChart(GrapeCity.Documents.Excel.Drawing.ChartType.ColumnStacked100, 250, 20, 360, 230);
            worksheet.Range["A1:B6"].Value = new object[,] {
                { 1, 5 },
                { 2, 4 },
                { 3, 3 },
                { 4, 2 },
                { 5, 1 },
                { 5, 3 },
            };

            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:B6"], GrapeCity.Documents.Excel.Drawing.RowCol.Columns);
            shape.Chart.ChartTitle.Text = "Column Stacked 100 Chart";
            
        }
    }
}
