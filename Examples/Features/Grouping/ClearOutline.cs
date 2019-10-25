using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Grouping
{
    public class ClearOutline : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            //1:20 rows' outline level will be 2.
            worksheet.Range["1:20"].Group();
            //1:10 rows' outline level will be 3.
            worksheet.Range["1:10"].Group();

            //1:20 rows' outline level will be 1.
            worksheet.Range["1:20"].ClearOutline();
        }
    }

}
