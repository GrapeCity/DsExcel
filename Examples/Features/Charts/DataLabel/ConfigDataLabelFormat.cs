using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Charts.DataLabel
{
    public class ConfigDataLabelFormat : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            GrapeCity.Documents.Excel.Drawing.IShape shape = worksheet.Shapes.AddChart(GrapeCity.Documents.Excel.Drawing.ChartType.ColumnClustered, 250, 20, 360, 230);
            worksheet.Range["A1:B5"].Value = new object[,]
            {
                {null, "S1"},
                {"Item1", -20},
                {"Item2", 30},
                {"Item3", 50 },
                {"Item3", 40 }
            };
            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:B5"], GrapeCity.Documents.Excel.Drawing.RowCol.Columns, true, true);
            GrapeCity.Documents.Excel.Drawing.ISeries series1 = shape.Chart.SeriesCollection[0];
            series1.HasDataLabels = true;
            series1.DataLabels.ShowSeriesName = true;

            //set series1's all data label's format.
            series1.DataLabels.Format.Fill.Color.RGB = Color.Pink;
            series1.DataLabels.Format.Line.Color.RGB = Color.Green;
            series1.DataLabels.Format.Line.Weight = 1;

            //set series1's specific data label's format.
            series1.DataLabels[2].Format.Fill.Color.RGB = Color.LightGreen;
            series1.Points[2].DataLabel.Format.Line.Color.RGB = Color.Gray;
            series1.Points[2].DataLabel.Format.Line.Weight = 2;

        }
    }
}
