using GrapeCity.Documents.Excel.Drawing;
using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Charts.NewCharts
{
    public class AddWaterfallChart : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            worksheet.Range["A1:B8"].Value = new object[,]
            {
                {"Starting Amt", 130},
                {"Measurement 1", 25},
                {"Measurement 2", -75},
                {"Subtotal", 80},
                {"Measurement 3", 45},
                {"Measurement 4", -65},
                {"Measurement 5", 80},
                {"Total", 140}
            };
            worksheet.Range["A:A"].Columns.AutoFit();

            //Create a waterfall chart.
            IShape shape = worksheet.Shapes.AddChart(ChartType.Waterfall, 300, 20, 300, 250);
            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:B8"]);

            //Set subtotal&total points.
            IPoints points = shape.Chart.SeriesCollection[0].Points;
            points[3].IsTotal = true;
            points[7].IsTotal = true;

            //Connector lines are not shown.
            ISeries series = shape.Chart.SeriesCollection[0];
            series.ShowConnectorLines = false;

            //Modify the fill color of the first legend entry.
            ILegendEntries legendEntries = shape.Chart.Legend.LegendEntries;
            legendEntries[0].Format.Fill.Color.ObjectThemeColor = ThemeColor.Accent6;
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
