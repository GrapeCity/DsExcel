using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Formatting.Alignment
{
    public class RotateCellContents : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            IRange rangeB2 = worksheet.Range["B2"];
            rangeB2.Value = "Rotated Cell Contents";
            rangeB2.HorizontalAlignment = HorizontalAlignment.Center;
            rangeB2.VerticalAlignment = VerticalAlignment.Center;
            //Rotate cell contents.
            rangeB2.Orientation = 15;
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
