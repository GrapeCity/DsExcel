using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Formatting.Fill
{
    public class PatternFill : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Range["A1"].Interior.Pattern = GrapeCity.Documents.Excel.Pattern.LightDown;
            worksheet.Range["A1"].Interior.Color = Color.Pink;
            worksheet.Range["A1"].Interior.PatternColorIndex = 5;
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
