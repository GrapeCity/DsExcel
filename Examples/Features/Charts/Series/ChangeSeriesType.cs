using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Charts.Series
{
    public class ChangeSeriesType : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            GrapeCity.Documents.Excel.Drawing.IShape shape = worksheet.Shapes.AddChart(GrapeCity.Documents.Excel.Drawing.ChartType.ColumnClustered, 250, 20, 360, 230);
            worksheet.Range["A1:D6"].Value = new object[,]
            {
                {null, "S1", "S2", "S3"},
                {"Item1", 10, 25, 25},
                {"Item2", -51, -36, 27},
                {"Item3", 52, -85, -30},
                {"Item4", 22, 65, 65},
                {"Item5", 23, 69, 69}
            };

            //Add series
            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:D6"], GrapeCity.Documents.Excel.Drawing.RowCol.Columns, true, true);
            GrapeCity.Documents.Excel.Drawing.ISeries series1 = shape.Chart.SeriesCollection[0];
            GrapeCity.Documents.Excel.Drawing.ISeries series2 = shape.Chart.SeriesCollection[1];
            //change series2's chart type.
            series2.ChartType = GrapeCity.Documents.Excel.Drawing.ChartType.Line;
            series2.Format.Line.Weight = 2;

        }
    }
}
