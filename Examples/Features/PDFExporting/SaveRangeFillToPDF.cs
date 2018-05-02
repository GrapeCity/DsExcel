using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting
{
    public class SaveRangeFillToPDF : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            IRange rangeA1B2 = worksheet.Range["A1:B2"];
            rangeA1B2.Merge();
            rangeA1B2.Interior.Pattern = GrapeCity.Documents.Excel.Pattern.LinearGradient;
            (rangeA1B2.Interior.Gradient as ILinearGradient).ColorStops[0].Color = Color.Red;
            (rangeA1B2.Interior.Gradient as ILinearGradient).ColorStops[1].Color = Color.Yellow;
            (rangeA1B2.Interior.Gradient as ILinearGradient).Degree = 90;

            IRange rangeE1E2 = worksheet.Range["D1:E2"];
            rangeE1E2.Merge();
            rangeE1E2.Interior.Pattern = GrapeCity.Documents.Excel.Pattern.LightDown;
            rangeE1E2.Interior.Color = Color.Pink;
            rangeE1E2.Interior.PatternColorIndex = 5;

            IRange rangeG1H2 = worksheet.Range["G1:H2"];
            rangeG1H2.Merge();
            rangeG1H2.Interior.Pattern = GrapeCity.Documents.Excel.Pattern.RectangularGradient;
            (rangeG1H2.Interior.Gradient as IRectangularGradient).ColorStops[0].Color = Color.Red;
            (rangeG1H2.Interior.Gradient as IRectangularGradient).ColorStops[1].Color = Color.Green;

            (rangeG1H2.Interior.Gradient as IRectangularGradient).Bottom = 0.2;
            (rangeG1H2.Interior.Gradient as IRectangularGradient).Right = 0.3;
            (rangeG1H2.Interior.Gradient as IRectangularGradient).Top = 0.4;
            (rangeG1H2.Interior.Gradient as IRectangularGradient).Left = 0.5;

            worksheet.Range["J1:K2"].Merge();
            worksheet.Range["J1:K2"].Interior.Color = Color.Green;
        }

        public override bool SavePdf
        {
            get
            {
                return true;
            }
        }

        public override bool ShowViewer
        {
            get
            {
                return false;
            }
        }

        public override bool IsNew
        {
            get
            {
                return true;
            }
        }
    }
}
