using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Formatting.NumberFormat
{
    public class CustomNumberFormat : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            //Set range values.
            worksheet.Range["A2:B2"].Value = -15.50;
            worksheet.Range["A3:B3"].Value = 555;
            worksheet.Range["A4:B4"].Value = 0;
            worksheet.Range["A5:B5"].Value = "Name";

            //Apply custom number format.
            worksheet.Range["B2:B5"].NumberFormat = "[Green]#.00;[Red]#.00;[Blue]0.00;[Cyan]\"product: \"@";
        }
    }
}
