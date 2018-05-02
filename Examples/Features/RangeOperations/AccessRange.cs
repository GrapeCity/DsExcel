using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.RangeOperations
{
    public class AccessRange : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            //use index to access range, Range A1.
            worksheet.Range[0, 0].Interior.Color = Color.LightGreen;
            worksheet.Range[0, 0, 2, 2].Value = 5;

            //use string to access range.
            worksheet.Range["A2"].Interior.Color = Color.LightYellow;
            worksheet.Range["C3:D4"].Interior.Color = Color.Tomato;
            worksheet.Range["A5:B7, C3, H5:N6"].Value = 2;

            //use Cells to access range.
            worksheet.Cells[5].Interior.Color = Color.LightBlue;
        }
    }
}
