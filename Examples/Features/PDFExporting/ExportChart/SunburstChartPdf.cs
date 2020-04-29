using GrapeCity.Documents.Excel.Drawing;
using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting.ExportChart
{
    public class SunburstChartPdf : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            worksheet.Range["A1:D16"].Value = new object[,]
            {
                {"Quarter", "Month", "Week", "Output"},
                {"1st", "Jan", null, 3.5},
                {null, "Feb", "Week1", 1.2},
                {null, null, "Week2", 0.8},
                {null, null, "Week3", 0.6},
                {null, null, "Week4", 0.5},
                {null, "Mar", null, 1.7 },
                {"2st", "Apr", null, 1.1},
                {null, "May", null, 0.8},
                {null, "Jun", null, 0.3},
                {"3st", "July", null, 0.7},
                {null, "Aug", null, 0.6},
                {null, "Sept", null, 0.1},
                {"4st", "Oct", null, 0.5},
                {null, "Nov", null, 0.4},
                {null, "Dec", null, 0.3},
            };

            //Create a sunburst chart.
            IShape shape = worksheet.Shapes.AddChart(Drawing.ChartType.Sunburst, 20, 320, 200, 200);
            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:D16"]);

            //Modify chart title text.
            shape.Chart.ChartTitle.Text = "Annual Report";
        }

        public override bool SavePdf
        {
            get
            {
                return true;
            }
        }

        public override bool ShowViewer
        {
            get
            {
                return false;
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
