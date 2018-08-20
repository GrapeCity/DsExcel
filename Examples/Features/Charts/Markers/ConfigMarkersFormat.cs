using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Charts.Markers
{
   public class ConfigMarkersFormat :ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            GrapeCity.Documents.Excel.Drawing.IShape shape = worksheet.Shapes.AddChart(GrapeCity.Documents.Excel.Drawing.ChartType.LineMarkers, 250, 20, 360, 230);
            worksheet.Range["A1:B6"].Value = new object[,]
            {
                {null, "S1"},
                {"Item1", 10},
                {"Item2", -51},
                {"Item3", 52},
                {"Item4", 22},
                {"Item5", 40}
            };

            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:B6"], GrapeCity.Documents.Excel.Drawing.RowCol.Columns, true, true);
            GrapeCity.Documents.Excel.Drawing.ISeries series1 = shape.Chart.SeriesCollection[0];
            //config line markers style
            series1.MarkerFormat.Fill.Color.RGB = Color.CornflowerBlue;
            series1.MarkerFormat.Line.Style = GrapeCity.Documents.Excel.Drawing.LineStyle.ThickThin;
            series1.MarkerFormat.Line.Color.RGB = Color.LightGreen;
            series1.MarkerFormat.Line.Weight = 3;

        }
    }
}
