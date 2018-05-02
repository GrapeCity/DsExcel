using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Formatting.Alignment
{
    public class WrapText : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            IRange rangeB3 = worksheet.Range["B3"];
            rangeB3.Value = "The WrapText property is applied to wrap the text within a cell";
            rangeB3.WrapText = true;

            worksheet.Rows[2].RowHeight = 150;
        }
    }
}
