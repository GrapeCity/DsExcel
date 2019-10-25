using GrapeCity.Documents.Excel.Drawing;
using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting.ExportShape
{
    public class ShapeWithGradientAndPatternFill : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet sheet = workbook.Worksheets[0];

            // Add a sun
            IShape sun = sheet.Shapes.AddShape(AutoShapeType.Sun, 20, 15, 200, 200);

            // Set gradient fill for sun
            sun.Fill.OneColorGradient(GradientStyle.Horizontal, 1, 0.8);
            sun.Fill.GradientStops.Insert(0Xff6347, 0);
            sun.Fill.GradientStops.Insert(0xFFA500, 0.2);
            sun.Fill.GradientStops.Insert(0xFFFF00, 0.4);
            sun.Fill.GradientStops.Insert(0xFF0000, 0.6);
            sun.Fill.GradientStops.Insert(0x0000FF, 0.8);
            sun.Fill.GradientStops.Insert(0x800080, 1);

            // Add a five point star
            IShape star = sheet.Shapes.AddShape(AutoShapeType.Shape5pointStar, 250, 15, 200, 200);

            // Set pattern fill for star
            star.Fill.Patterned(PatternType.DottedDiamond);
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
    }
}
