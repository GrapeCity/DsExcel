using GrapeCity.Documents.Excel.Drawing;
using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting.ExportShape
{
    public class ShapeWithLine : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet sheet = workbook.Worksheets[0];

            IShape rectangle = sheet.Shapes.AddShape(AutoShapeType.Rectangle, 20, 50, 200, 200);
            rectangle.Line.DashStyle = LineDashStyle.Dash;
            rectangle.Line.Style = LineStyle.Single;
            rectangle.Line.Weight = 7;
            rectangle.Line.Color.RGB = System.Drawing.Color.Yellow;

            IShape donut = sheet.Shapes.AddShape(AutoShapeType.Donut, 260, 50, 200, 200);
            donut.Line.DashStyle = LineDashStyle.DashDotDot;
            donut.Line.Style = LineStyle.Single;
            donut.Line.Weight = 7;
            donut.Line.Color.RGB = System.Drawing.Color.Red;
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
