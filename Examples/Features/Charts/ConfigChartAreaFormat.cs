using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.Charts
{
    public class ConfigChartAreaFormat : ExampleBase
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

            GrapeCity.Documents.Spread.Drawing.IChartArea chartarea = shape.Chart.ChartArea;

            //Format
            chartarea.Format.Fill.Color.RGB = Color.White;
            chartarea.Format.Line.Color.RGB = Color.LightGreen;
            chartarea.Format.Line.Weight = 2;

            //3d format just take effort in chart area.
            chartarea.Format.ThreeD.RotationX = 60;
            chartarea.Format.ThreeD.RotationY = 20;
            chartarea.Format.ThreeD.RotationZ = 100;
            chartarea.Format.ThreeD.Z = 20;
            chartarea.Format.ThreeD.Perspective = 20;
            chartarea.Format.ThreeD.Depth = 5;

            //Font
            chartarea.Font.Bold = true;
            chartarea.Font.Italic = true;
            chartarea.Font.Color.RGB = Color.LightGreen;
            //rounded corners.
            chartarea.RoundedCorners = true;
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
