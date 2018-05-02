using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Charts.Axes
{
    public class ConfigTickMarkLabel : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            GrapeCity.Documents.Excel.Drawing.IShape shape = worksheet.Shapes.AddChart(GrapeCity.Documents.Excel.Drawing.ChartType.ColumnClustered, 250, 20, 360, 230);
            worksheet.Range["A1:D6"].Value = new object[,]
            {
                {null, "S1", "S2", "S3"},
                {1, -25, 25, 25},
                {2, 51, 36, 27},
                {3, 52, 80, 30},
                {4, 22, -20, 65},
                {5, 23, 69, 69}
            };
            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:D6"], GrapeCity.Documents.Excel.Drawing.RowCol.Columns, true, true);
            GrapeCity.Documents.Excel.Drawing.IAxis category_axis = shape.Chart.Axes.Item(GrapeCity.Documents.Excel.Drawing.AxisType.Category);
            GrapeCity.Documents.Excel.Drawing.IAxis value_axis = shape.Chart.Axes.Item(GrapeCity.Documents.Excel.Drawing.AxisType.Value);

            //config tick label's format
            category_axis.TickLabelPosition = GrapeCity.Documents.Excel.Drawing.TickLabelPosition.NextToAxis;
            category_axis.TickLabelSpacing = 2;
            category_axis.TickLabels.Font.Color.RGB = Color.DarkOrange;           
            category_axis.TickLabels.Font.Size = 12;
            category_axis.TickLabels.NumberFormat = "#,##0.00";
            value_axis.TickLabels.NumberFormat = "#,##0;[Red]#,##0";
        }
    }
}
