using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.RangeOperations
{
    public class GetRangeRowCountColumnCount : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            var range = worksheet.Range["A5:B7"];

            //cell count is 6.
            var cellcount = range.Count;
            //cell count is 6.
            var cellcount1 = range.Cells.Count;
            //row count is 3.
            var rowcount = range.Rows.Count;
            //column count is 2.
            var columncount = range.Columns.Count;

        }

        public override bool CanDownload
        {
            get
            {
                return false;
            }
        }

        public override bool ShowViewer
        {
            get
            {
                return false;
            }
        }
    }
}
