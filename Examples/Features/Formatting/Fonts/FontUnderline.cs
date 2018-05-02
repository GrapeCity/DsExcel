using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Formatting.Fonts
{
    public class FontUnderline : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            worksheet.Range["A1"].Value = "Single Underline";
            worksheet.Range["A1"].Font.Underline = UnderlineType.Single;
        }
    }
}
