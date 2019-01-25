﻿using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Charts
{
    public class ChartCopy : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            //Create chart, chart's range is Range["G1:M21"]
            GrapeCity.Documents.Excel.Drawing.IShape shape = worksheet.Shapes.AddChart(GrapeCity.Documents.Excel.Drawing.ChartType.ColumnClustered, 300, 10, 300, 300);
            worksheet.Range["A1:D6"].Value = new object[,]
            {
               {null, "S1", "S2", "S3"},
               {"Item1", 10, 25, 25},
               {"Item2", -51, -36, 27},
               {"Item3", 52, -85, -30},
               {"Item4", 22, 65, 65},
               {"Item5", 23, 69, 69}
            };

            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:D6"], GrapeCity.Documents.Excel.Drawing.RowCol.Columns, true, true);

            //Range["G1:M21"] must contain chart's range, copy a new shape to Range["N1:T21"]
            worksheet.Range["G1:M21"].Copy(worksheet.Range["N1"]);
            //worksheet.Range["G1:M21"].Copy(worksheet.Range["N1:T21"]);

            //Cross sheet copy, copy a new chart to worksheet2's Range["N1:T21"]
            //IWorksheet worksheet2 = workbook.Worksheets.Add()
            //worksheet.Range["G1:M21"].Copy(worksheet2.Range["E1"]);
            //worksheet.Range["G1:M21"].Copy(worksheet2.Range["N1:T21"]);

        }
    }
}
