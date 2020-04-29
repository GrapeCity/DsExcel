using GrapeCity.Documents.Excel.Drawing;
using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Charts.ErrorBars
{
    public class ConfigXYErrorBars : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            GrapeCity.Documents.Excel.Drawing.IShape shape = worksheet.Shapes.AddChart(GrapeCity.Documents.Excel.Drawing.ChartType.XYScatter, 250, 20, 360, 230);
            worksheet.Range["A1:D7"].Value = new object[,] {
                { "Blue", null, "Red", null },
                { 55, 964, 67, 475 },
                { 20, 825, 10, 163 },
                { 77, 840, 87, 224 },
                { 182, 596, 46, 196 },
                { 190, 384, 100, 377 },
                { 140, 503, 92, 47 },
            };
            shape.Chart.SeriesCollection.Add(worksheet.Range["A2:B7"], RowCol.Columns);
            shape.Chart.SeriesCollection.Add(worksheet.Range["C2:D7"], RowCol.Columns);
            shape.Chart.ChartTitle.Text = "Scatter Chart";

            // Get first series
            ISeries series1 = shape.Chart.SeriesCollection[0];

            // Set HasErrorBars as true
            series1.HasErrorBars = true;

            // Config y-direction error bar
            series1.YErrorBar.ValueType = ErrorBarType.FixedValue;
            series1.YErrorBar.Amount = 500;

            // Config x-direction error bar
            series1.XErrorBar.ValueType = ErrorBarType.FixedValue;
            series1.XErrorBar.Amount = 20;
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
