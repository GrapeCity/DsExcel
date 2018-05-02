using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Formatting.Fonts
{
    public class FontName : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            //use Name property to set font name.
            worksheet.Range["A1"].Value = "Calibri";
            worksheet.Range["A1"].Font.Name = "Calibri";

            //use ThemeFont property to set font name.
            worksheet.Range["A2"].Value = "Major theme font";
            worksheet.Range["A2"].Font.ThemeFont = ThemeFont.Major;
        }
    }
}
