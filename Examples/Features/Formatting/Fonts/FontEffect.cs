using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Formatting.Fonts
{
    public class FontEffect : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            worksheet.Range["A1"].Value = "Strikethrough";
            worksheet.Range["A1"].Font.Strikethrough = true;

            worksheet.Range["A2"].Value = "Superscript";
            worksheet.Range["A2"].Font.Superscript = true;

            worksheet.Range["A3"].Value = "Subscript";
            worksheet.Range["A3"].Font.Subscript = true;
        }

        public override bool ShowViewer
        {
            get
            {
                return false;
            }
        }

        public override bool ShowScreenshot
        {
            get
            {
                return true;
            }
        }

    }
}
