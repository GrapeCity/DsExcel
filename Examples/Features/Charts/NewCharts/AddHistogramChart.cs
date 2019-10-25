using GrapeCity.Documents.Excel.Drawing;
using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Charts.NewCharts
{
    public class AddHistogramChart : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            worksheet.Range["A1:B11"].Value = new object[,]
            {
                {"Complaint", "Count"},
                {"Too noisy", 27},
                {"Overpriced", 789},
                {"Food is tasteless", 65},
                {"Food is not fresh", 9},
                {"Food is too salty", 15},
                {"Not clean", 30},
                {"Unfriendly staff", 12},
                {"Wait time", 109},
                { "No atmosphere", 45},
                {"Small portions", 621 }
            };
            worksheet.Range["A:A"].Columns.AutoFit();

            //Create a histogram chart.
            IShape shape = worksheet.Shapes.AddChart(ChartType.Histogram, 300, 20, 300, 200);
            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:B11"]);

            //Sets bins type by count.
            shape.Chart.ChartGroups[0].BinsType = BinsType.BinsTypeBinCount;
            shape.Chart.ChartGroups[0].BinsCountValue = 3;

            //Set overflow bin value
            shape.Chart.ChartGroups[0].BinsOverflowEnabled = true;
            shape.Chart.ChartGroups[0].BinsOverflowValue = 500;
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

        public override bool IsNew
        {
            get
            {
                return true;
            }
        }
    }
}
