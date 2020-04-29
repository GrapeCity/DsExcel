using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting.ExportChart
{
    public class BarChartPdf : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            GrapeCity.Documents.Excel.Drawing.IShape shape = worksheet.Shapes.AddChart(GrapeCity.Documents.Excel.Drawing.ChartType.BarClustered, 20, 20, 360, 230);
            worksheet.Range["A20:D21"].Value = new object[,] {
                { 100,200,300,400 },
                { 100,200,300,400 },
            };
            shape.Chart.SeriesCollection.Add(worksheet.Range["A20:D21"], GrapeCity.Documents.Excel.Drawing.RowCol.Columns);
            shape.Chart.ChartTitle.Text = "Bar Clustered Chart";
            shape.Chart.Legend.Position = GrapeCity.Documents.Excel.Drawing.LegendPosition.Left;
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
    }
}
