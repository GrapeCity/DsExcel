using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.Charts.Series
{
    public class ConfigDataLabelTextAndFormat : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            GrapeCity.Documents.Spread.Drawing.IShape shape = worksheet.Shapes.AddChart(GrapeCity.Documents.Spread.Drawing.ChartType.ColumnClustered, 300, 10, 300, 300);
            worksheet.Range["A1:D6"].Value = new object[,]
            {
                {null, "S1", "S2", "S3"},
                {"Item1", 10, 25, 25},
                {"Item2", -51, -36, 27},
                {"Item3", 52, -85, -30},
                {"Item4", 22, 65, 65},
                {"Item5", 23, 69, 69}
            };
            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:D6"], GrapeCity.Documents.Spread.Drawing.RowCol.Columns, true, true);

            GrapeCity.Documents.Spread.Drawing.ISeries series1 = shape.Chart.SeriesCollection[0];
            series1.HasDataLabels = true;

            //set series1's all data label's format.
            series1.DataLabels.Format.Fill.Color.RGB = Color.Pink;
            series1.DataLabels.Format.Line.Color.RGB = Color.Green;
            series1.DataLabels.Format.Line.Weight = 1;

            //set series1's specific data label's format.
            series1.DataLabels[2].Format.Fill.Color.RGB = Color.LightGreen;
            series1.Points[2].DataLabel.Format.Line.Color.RGB = Color.Pink;
            series1.Points[2].DataLabel.Format.Line.Weight = 2;

            //customize data lables' text.
            series1.DataLabels.ShowCategoryName = true;
            series1.DataLabels.ShowSeriesName = true;
            series1.DataLabels.ShowLegendKey = true;
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
