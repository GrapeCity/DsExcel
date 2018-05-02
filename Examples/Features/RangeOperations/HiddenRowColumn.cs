using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.RangeOperations
{
    public class HiddenRowColumn : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Range["E1"].Value = 1;

            //Hidden row 2:6.
            worksheet.Range["2:6"].Hidden = true;

            //Hidden column A:D.
            worksheet.Range["A:D"].Hidden = true;
        }
    }
}
