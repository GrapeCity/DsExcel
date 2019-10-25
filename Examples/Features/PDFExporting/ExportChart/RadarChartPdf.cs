using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting.ExportChart
{
    public class RadarChartPdf : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            GrapeCity.Documents.Excel.Drawing.IShape shape = worksheet.Shapes.AddChart(GrapeCity.Documents.Excel.Drawing.ChartType.Radar, 20, 20, 360, 230);
            worksheet.Range["B20:C20"].Value = new string[,] { { "S1", "S2" } };
            worksheet.Range["A21:A25"].Value = new string[,] { { "A" }, { "B" }, { "C" }, { "D" }, { "E" } };
            worksheet.Range["B21:C25"].Value = new double[,] { { 10, 25 }, { 51, 36 }, { 52, 85 }, { 22, 65 }, { 23, 69 } };

            shape.Chart.SeriesCollection.Add(worksheet.Range["A20:C25"]);
        }

        public override bool IsNew
        {
            get
            {
                return true;
            }
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
