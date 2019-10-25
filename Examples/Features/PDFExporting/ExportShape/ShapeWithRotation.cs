using GrapeCity.Documents.Excel.Drawing;
using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting.ExportShape
{
    public class ShapeWithRotation : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet sheet = workbook.Worksheets[0];

            sheet.Range["D2"].Value = "rectangle with 30 degrees";
            sheet.Range["I2"].Value = "right arrow with 40 degrees";

            // Add a rectangle with rotation
            IShape rectangle = sheet.Shapes.AddShape(AutoShapeType.Rectangle, 50, 50, 200, 50);
            rectangle.Rotation = 30;

            // Add a right arrow with rotation
            IShape rightArrowWithRotation = sheet.Shapes.AddShape(AutoShapeType.RightArrow, 270, 50, 200, 100);
            rightArrowWithRotation.Rotation = 40;
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
