using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Formatting.Fill
{
    public class RectangularGradientFill : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Range["A1"].Interior.Pattern = GrapeCity.Documents.Excel.Pattern.RectangularGradient;
            (worksheet.Range["A1"].Interior.Gradient as IRectangularGradient).ColorStops[0].Color = Color.Red;
            (worksheet.Range["A1"].Interior.Gradient as IRectangularGradient).ColorStops[1].Color = Color.Green;

            (worksheet.Range["A1"].Interior.Gradient as IRectangularGradient).Bottom = 0.2;
            (worksheet.Range["A1"].Interior.Gradient as IRectangularGradient).Right = 0.3;
            (worksheet.Range["A1"].Interior.Gradient as IRectangularGradient).Top = 0.4;
            (worksheet.Range["A1"].Interior.Gradient as IRectangularGradient).Left = 0.5;
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
