using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Formatting.Alignment
{
    public class HAlignVAlign : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            worksheet.Columns[0].ColumnWidth = 17;

            IRange rangeA1 = worksheet.Range["A1"];
            rangeA1.Value = "Right and top";
            rangeA1.HorizontalAlignment = HorizontalAlignment.Right;
            rangeA1.VerticalAlignment = VerticalAlignment.Top;

            IRange rangeA2 = worksheet.Range["A2"];
            rangeA2.Value = "Center";
            rangeA2.HorizontalAlignment = HorizontalAlignment.Center;
            rangeA2.VerticalAlignment = VerticalAlignment.Center;

            IRange rangeA3 = worksheet.Range["A3"];
            rangeA3.Value = "Left and bottom, indent";
            rangeA3.IndentLevel = 1;
        }
    }
}
