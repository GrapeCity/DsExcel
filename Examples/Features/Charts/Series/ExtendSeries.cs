using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Charts.Series
{
    public class ExtendSeries : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            GrapeCity.Documents.Excel.Drawing.IShape shape = worksheet.Shapes.AddChart(GrapeCity.Documents.Excel.Drawing.ChartType.ColumnClustered, 250, 20, 360, 230);
            worksheet.Range["A1:D4"].Value = new object[,]
            {
                {null, "S1", "S2", "S3"},
                {"Item1", 10, 25, 50},
                {"Item2", 15, -36, 40},
                {"Item3", 52, 40, -30}, 
            };
            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:D4"], GrapeCity.Documents.Excel.Drawing.RowCol.Columns, true, true);

            worksheet.Range["A12:D13"].Value = new object[,]
            {
                {"Item5", 10, 20, -30},
                {"Item6", 20, 40, 80},
            };

            //add new data point to existing series.
            shape.Chart.SeriesCollection.Extend(worksheet.Range["A12:D13"], GrapeCity.Documents.Excel.Drawing.RowCol.Columns, true);
        }
    }
}
