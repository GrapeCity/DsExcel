using GrapeCity.Documents.Excel.Drawing;
using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Charts.NewCharts
{
    public class AddParetoChart : ExampleBase
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

            //Create a pareto chart.
            IShape shape = worksheet.Shapes.AddChart(ChartType.Pareto, 300, 20, 300, 200);
            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:B11"]);

            //Set bins type by size.
            shape.Chart.ChartGroups[0].BinsType = BinsType.BinsTypeBinSize;
            shape.Chart.ChartGroups[0].BinWidthValue = 300;

            //Set underflow bin value.
            shape.Chart.ChartGroups[0].BinsUnderflowEnabled = true;
            shape.Chart.ChartGroups[0].BinsUnderflowValue = 50;
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
