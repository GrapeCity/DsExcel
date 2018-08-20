using System;
using System.Collections.Generic;
using System.Drawing;
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
            range.Cells[4].Value = "A7";

            //cell is B6
            range.Cells[1, 1].Value = "B6";

            //row count is 3 and range is A6:B6.
            var rowCount = range.Rows.Count;
            var row = range.Rows[1].ToString();

            //set interior color for row range A6:B6.
            range.Rows[1].Interior.Color = Color.LightBlue;

            //column count is 2 and range is B5:B7.
            var columnCount = range.Columns.Count;
            var column = range.Columns[1].ToString();

            //set values for column range B5:B7.
            range.Columns[1].Interior.Color = Color.LightSkyBlue;

            //entire rows are from row 5 to row 7
            var entirerow = range.EntireRow.ToString();

            //entire columns are from column A to column B
            var entireColumn = range.EntireColumn.ToString();

        }

        public override bool IsUpdate
        {
            get
            {
                return true;
            }
        }
    }
}
