using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.Charts.Series
{
    public class CreateTrendline : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            GrapeCity.Documents.Spread.Drawing.IShape shape = worksheet.Shapes.AddChart(GrapeCity.Documents.Spread.Drawing.ChartType.ColumnClustered, 300, 10, 300, 300);
            worksheet.Range["A1:D6"].Value = new object[,]
            {
                {null, "S1", "S2", "S3"},
                {"Item1", 10, 25, 25},
                {"Item2", -51, -36, 27},
                {"Item3", 52, -85, -30},
                {"Item4", 22, 65, 65},
                {"Item5", 23, 69, 69}
            };
            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:D6"], GrapeCity.Documents.Spread.Drawing.RowCol.Columns, true, true);

            GrapeCity.Documents.Spread.Drawing.ISeries series1 = shape.Chart.SeriesCollection[0];

            series1.Trendlines.Add();
            series1.Trendlines[0].Type = GrapeCity.Documents.Spread.Drawing.TrendlineType.Linear;
            series1.Trendlines[0].Forward = 1;
            series1.Trendlines[0].Backward = 0.5;
            series1.Trendlines[0].Intercept = 2.5;
            series1.Trendlines[0].DisplayEquation = true;
            series1.Trendlines[0].DisplayRSquared = true;
        }

        public override bool ShowViewer
        {
            get
            {
                return false;
            }
        }
    }
}
