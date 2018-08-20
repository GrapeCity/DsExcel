using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Charts.ChartTitle
{
    public class ConfigChartTitleRichText :ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            GrapeCity.Documents.Excel.Drawing.IShape shape = worksheet.Shapes.AddChart(GrapeCity.Documents.Excel.Drawing.ChartType.ColumnClustered, 250, 20, 360, 230);
            worksheet.Range["A1:D6"].Value = new object[,]
            {
                {null, "S1", "S2", "S3"},
                {"Item1", 10, 25, 25},
                {"Item2", -20, 36, 27},
                {"Item3", 62, 70, -30},
                {"Item4", 22, 65, 65},
                {"Item5", 23, 50, 50}
            };
            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:D6"], GrapeCity.Documents.Excel.Drawing.RowCol.Columns, true, true);

            //config chart title style with rich text
            shape.Chart.HasTitle = true;
            shape.Chart.ChartTitle.TextFrame.TextRange.Paragraphs.Add("ChartSubtitle");
            shape.Chart.ChartTitle.TextFrame.TextRange.Paragraphs.Add("ChartTitle", 0);
            shape.Chart.ChartTitle.TextFrame.TextRange.Paragraphs[0].Font.Color.RGB = Color.CornflowerBlue;
            shape.Chart.ChartTitle.TextFrame.TextRange.Paragraphs[0].Font.Size = 15;
            shape.Chart.ChartTitle.TextFrame.TextRange.Paragraphs[1].Font.Color.RGB = Color.Orange;
            shape.Chart.ChartTitle.TextFrame.TextRange.Paragraphs[1].Font.Size = 10;

        }

        public override bool ShowViewer
        {
            get
            {
                return false;
            }
        }

        public override bool ShowScreenshot
        {
            get
            {
                return true;
            }
        }

    }
}
