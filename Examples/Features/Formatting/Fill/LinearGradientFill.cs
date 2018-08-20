using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Formatting.Fill
{
    public class LinearGradientFill : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Range["A1"].Interior.Pattern = GrapeCity.Documents.Excel.Pattern.LinearGradient;
            (worksheet.Range["A1"].Interior.Gradient as ILinearGradient).ColorStops[0].Color = Color.Red;
            (worksheet.Range["A1"].Interior.Gradient as ILinearGradient).ColorStops[1].Color = Color.Yellow;

            (worksheet.Range["A1"].Interior.Gradient as ILinearGradient).Degree = 90;
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
