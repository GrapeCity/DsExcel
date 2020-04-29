using GrapeCity.Documents.Excel.Drawing;
using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Charts.ErrorBars
{
    public class ConfigCustomErrorBar : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            GrapeCity.Documents.Excel.Drawing.IShape shape = worksheet.Shapes.AddChart(ChartType.Line, 250, 20, 360, 230);
            worksheet.Range["A1:D4"].Value = new object[,]
            {
                {null, "Q1", "Q2", "Q3"},
                {"Mobile Phones", 1330, 2330, 3330},
                {"Laptops", 4032, 5632, 6197},
                {"Tablets", 6233, 7233, 8233}
            };
            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:D4"], RowCol.Rows);

            // Get first series
            ISeries series1 = shape.Chart.SeriesCollection[0];

            // Set HasErrorBars as true
            series1.HasErrorBars = true;

            // Set value type of first series' error bar as custom
            series1.YErrorBar.ValueType = ErrorBarType.Custom;
            series1.YErrorBar.Plus = "={2000}";
            series1.YErrorBar.Minus = "={1000}";
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
