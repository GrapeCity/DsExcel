using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Formatting.Fonts
{
    public class FontStyle : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            worksheet.Range["A1"].Value = "Bold";
            worksheet.Range["A1"].Font.Bold = true;

            worksheet.Range["A2"].Value = "Italic";
            worksheet.Range["A2"].Font.Italic = true;

            worksheet.Range["A3"].Value = "Bold Italic";
            worksheet.Range["A3"].Font.Bold = true;
            worksheet.Range["A3"].Font.Italic = true;
        }
    }
}
