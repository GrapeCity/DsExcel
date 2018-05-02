using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.Charts.Axes
{
    public class ConfigAxesFormat : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            GrapeCity.Documents.Spread.Drawing.IShape shape = worksheet.Shapes.AddChart(GrapeCity.Documents.Spread.Drawing.ChartType.Line, 300, 10, 300, 300);
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

            GrapeCity.Documents.Spread.Drawing.IAxis category_axis = shape.Chart.Axes.Item(GrapeCity.Documents.Spread.Drawing.AxisType.Category);
            GrapeCity.Documents.Spread.Drawing.IAxis value_axis = shape.Chart.Axes.Item(GrapeCity.Documents.Spread.Drawing.AxisType.Value);

            //set category axis's format.
            category_axis.Format.Fill.Color.ObjectThemeColor = ThemeColor.Accent6;
            category_axis.Format.Fill.Color.TintAndShade = 0.8;
            category_axis.Format.Line.Color.RGB = Color.LightGreen;
            category_axis.Format.Line.Weight = 3;
            category_axis.Format.Line.Style = GrapeCity.Documents.Spread.Drawing.LineStyle.Single;

            //set value axis's format.
            value_axis.Format.Line.Color.RGB = Color.Pink;
            value_axis.Format.Line.Weight = 2;
            value_axis.Format.Line.Style = GrapeCity.Documents.Spread.Drawing.LineStyle.ThinThin;
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
