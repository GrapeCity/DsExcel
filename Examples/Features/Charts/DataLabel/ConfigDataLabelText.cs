using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Charts.DataLabel
{
    public class ConfigDataLabelText : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            GrapeCity.Documents.Excel.Drawing.IShape shape = worksheet.Shapes.AddChart(GrapeCity.Documents.Excel.Drawing.ChartType.ColumnClustered, 250, 20, 360, 230);
            worksheet.Range["A1:B5"].Value = new object[,]
            {
                {null, "S1","S2"},
                {"Item1", -20,10},
                {"Item2", 30,60},
                {"Item3", 50 ,80},
                {"Item3", 40,70 }
            };
            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:B5"], GrapeCity.Documents.Excel.Drawing.RowCol.Columns, true, true);
            GrapeCity.Documents.Excel.Drawing.ISeries series1 = shape.Chart.SeriesCollection[0];
            series1.HasDataLabels = true;
            
            //customize data lables' text.
            series1.DataLabels.ShowCategoryName = true;
            series1.DataLabels.ShowSeriesName = true;
            series1.DataLabels.ShowLegendKey = true;

        }
    }
}
