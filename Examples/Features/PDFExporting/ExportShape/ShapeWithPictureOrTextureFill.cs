using GrapeCity.Documents.Excel.Drawing;
using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting.ExportShape
{
    public class ShapeWithPictureOrTextureFill : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet sheet = workbook.Worksheets[0];

            // Get stream of picture
            System.IO.Stream stream = this.GetResourceStream("logo.png");

            // Add a rectangle
            IShape rectangle = sheet.Shapes.AddShape(AutoShapeType.Rectangle, 20, 20, 250, 50);
            // Set picture fill
            rectangle.Fill.UserPicture(stream, Drawing.ImageType.PNG);
            rectangle.Line.Transparency = 1;

            // Add a oval
            IShape oval = sheet.Shapes.AddShape(AutoShapeType.Oval, 20, 90, 250, 50);
            // Set picture fill
            oval.Fill.UserPicture(stream, Drawing.ImageType.PNG);
            oval.Line.Color.RGB = System.Drawing.Color.FromArgb(0x49129E);

            // Add a five point star
            IShape star = sheet.Shapes.AddShape(AutoShapeType.Shape5pointStar, 300, 20, 100, 100);
            // Set picture fill
            star.Fill.PresetTextured(Drawing.PresetTexture.WaterDroplets);
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

        public override string[] UsedResources
        {
            get
            {
                return new string[] { "logo.png" };
            }
        }
    }
}
