using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Formatting.Alignment
{
    public class TextOrientation : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            IRange rangeC1 = worksheet.Range["C1"];
            rangeC1.Value = "The ReadingOrder property is applied to set text direction.";
            rangeC1.ReadingOrder = ReadingOrder.RightToLeft;
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
