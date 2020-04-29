using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Worksheets
{
    public class AddChartSheet : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            worksheet.Range["A1:D6"].Value = new object[,]
            {
                {null, "S1", "S2", "S3"},
                {"Item1", 10, 25, 25},
                {"Item2", 51, 36, 27},
                {"Item3", 52, 85, 30},
                {"Item4", 22, 65, 65},
                {"Item5", 23, 69, 69}
            };

            //Add a chart sheet.
            IWorksheet chartSheet = workbook.Worksheets.Add(SheetType.Chart);

            //Add the main chart for the chart sheet.
            Drawing.IShape mainChart = chartSheet.Shapes.AddChart(Drawing.ChartType.ColumnClustered, 100, 100, 200, 200);
            mainChart.Chart.SeriesCollection.Add(worksheet.Range["A1:D6"]);

            //Make the chart sheet the active sheet.
            chartSheet.Activate();
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
