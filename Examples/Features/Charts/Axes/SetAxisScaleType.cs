using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Charts.Axes
{
    public class SetAxisScaleType : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            GrapeCity.Documents.Excel.Drawing.IShape shape = worksheet.Shapes.AddChart(GrapeCity.Documents.Excel.Drawing.ChartType.ColumnClustered, 250, 20, 360, 230);
            worksheet.Range["A1:D5"].Value = new object[,]
            {
                {null, "S1", "S2", "S3"},
                {"Item1", 4, 25, 7},
                {"Item2", 15, -10, 18},
                {"Item3", 45, 90, 20},
                {"Item4", 8, 20, 11},
            };
            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:D5"], GrapeCity.Documents.Excel.Drawing.RowCol.Columns, true, true);

            GrapeCity.Documents.Excel.Drawing.IAxis value_axis = shape.Chart.Axes.Item(GrapeCity.Documents.Excel.Drawing.AxisType.Value);
            value_axis.ScaleType = GrapeCity.Documents.Excel.Drawing.ScaleType.Logarithmic;
            value_axis.LogBase = 5;
        }
    }
}
