using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Charts.DataPoint
{
    public class ConfigSecondarySection :ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            GrapeCity.Documents.Excel.Drawing.IShape shape = worksheet.Shapes.AddChart(GrapeCity.Documents.Excel.Drawing.ChartType.PieOfPie, 250, 20, 360, 230);
            worksheet.Range["A1:D6"].Value = new object[,]
            {
                {null, "S1", "S2", "S3"},
                {"Item1", 10, 25, 25},
                {"Item2", -51, -36, 27},
                {"Item3", 52, -85, -30},
                {"Item4", 22, 65, 65},
                {"Item5", 23, 69, 69}
            };
            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:D6"], GrapeCity.Documents.Excel.Drawing.RowCol.Columns, true, true);
            GrapeCity.Documents.Excel.Drawing.ISeries series1 = shape.Chart.SeriesCollection[0];
            series1.HasDataLabels = true;

            //config secondary section for pie of pie chart
            shape.Chart.ChartGroups[0].SplitType = GrapeCity.Documents.Excel.Drawing.ChartSplitType.SplitByCustomSplit;
            series1.Points[0].SecondaryPlot = true;
            series1.Points[1].SecondaryPlot = false;
            series1.Points[2].SecondaryPlot = true;
            series1.Points[3].SecondaryPlot = false;
            series1.Points[4].SecondaryPlot = true;

        }
        public override bool ShowViewer
        {
            get
            {
                return false;
            }
        }

        public override bool ShowScreenshot
        {
            get
            {
                return true;
            }
        }
    }
}
