using GrapeCity.Documents.Excel.Drawing;
using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Charts.NewCharts
{
    public class AddFunnelChart : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            worksheet.Range["A1:B7"].Value = new object[,]
            {
                {"Stage", "Amount"},
                {"Prospects", 500},
                {"Qualified prospects", 425},
                {"Needs analysis", 200},
                {"Price quotes", 150},
                {"Negotiations", 100},
                {"Closed sales", 90}
            };
            worksheet.Range["A:A"].Columns.AutoFit();

            //Create a funnel chart.
            IShape shape = worksheet.Shapes.AddChart(ChartType.Funnel, 300, 20, 300, 200);
            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:B7"]);

            //Set the axis invisible.
            IAxis axis = shape.Chart.Axes.Item(AxisType.Category, AxisGroup.Primary);
            axis.Visible = false;            
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
