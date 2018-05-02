using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.RangeOperations
{
    public class AccessCellsRowsColumns : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            var range = worksheet.Range["A5:B7"];

            //set value for cell A7.
            var cell = range.Cells[4];
            cell.Value = 1;

            //set interior color for row range A6:B6.
            var row = range.Rows[1];
            row.Interior.Color = Color.LightBlue;

            //set values for column range B5:B7.
            var column = range.Columns[1];
            column.Value = 2;
        }
    }
}
