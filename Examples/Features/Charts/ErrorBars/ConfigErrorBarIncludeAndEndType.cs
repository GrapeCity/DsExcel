using GrapeCity.Documents.Excel.Drawing;
using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Charts.ErrorBars
{
    public class ConfigErrorBarIncludeAndEndType : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            GrapeCity.Documents.Excel.Drawing.IShape shape = worksheet.Shapes.AddChart(GrapeCity.Documents.Excel.Drawing.ChartType.Line, 250, 20, 360, 230);
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

            // Config first series' error bar
            series1.YErrorBar.Type = ErrorBarInclude.Both;
            series1.YErrorBar.EndStyle = EndStyleCap.Cap;

            // Get third series
            ISeries series2 = shape.Chart.SeriesCollection[2];

            // Set HasErrorBars as true
            series2.HasErrorBars = true;

            // Config third series' error bar
            series2.YErrorBar.Type = ErrorBarInclude.Plus;
            series2.YErrorBar.EndStyle = EndStyleCap.NoCap;
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
