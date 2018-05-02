using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.RangeOperations
{
    public class MergeCells : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            //A1:C4 is a single merged cell.
            worksheet.Range["A1:C4"].Merge();
            //H5:J5 is a single merged cell.
            //H6:J6 is a single merged cell.
            worksheet.Range["H5:J6"].Merge(true);

            //select A1:H5's entire merge area A1:J5, entire merge area is a bounding rectangle.
            var entireMergeArea = worksheet.Range["A1:H5"].EntireMergeArea;
            entireMergeArea.Select();

            //judge if H5 is a merged cell.
            if (worksheet.Range["J5"].MergeCells)
            {
                //set value to the top left cell of the merge area.
                worksheet.Range["J5"].EntireMergeArea[0, 0].Value = 1;
            }
        }
    }
}
