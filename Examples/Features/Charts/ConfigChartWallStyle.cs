using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.Charts
{
    public class ConfigChartWallStyle : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            GrapeCity.Documents.Spread.Drawing.IShape shape = worksheet.Shapes.AddChart(GrapeCity.Documents.Spread.Drawing.ChartType.Column3D, 300, 10, 300, 300);
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

            //config back wall and side wall together.
            shape.Chart.Walls.Thickness = 10;
            shape.Chart.Walls.Format.Fill.Color.RGB = Color.LightPink;
            shape.Chart.Walls.Format.Line.Color.RGB = Color.LightBlue;

            //config back wall individually.
            shape.Chart.BackWall.Thickness = 10;
            shape.Chart.BackWall.Format.Fill.Color.RGB = Color.LightGreen;
            shape.Chart.BackWall.Format.Line.Color.RGB = Color.LightBlue;
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
