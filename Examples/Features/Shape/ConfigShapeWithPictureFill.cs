using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;

namespace GrapeCity.Documents.Excel.Examples.Features.Shape
{
    public class ConfigShapeWithPictureFill : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            GrapeCity.Documents.Excel.Drawing.IShape shape = worksheet.Shapes.AddShape(GrapeCity.Documents.Excel.Drawing.AutoShapeType.Parallelogram, 20, 20, 200, 100);

            System.IO.Stream stream = this.GetResourceStream("logo.png");

            shape.Fill.UserPicture(stream, GrapeCity.Documents.Excel.Drawing.ImageType.JPG);
            shape.PictureFormat.ColorType = GrapeCity.Documents.Excel.Drawing.PictureColorType.Grayscale;
            shape.PictureFormat.Brightness = 0.6;
            shape.PictureFormat.Contrast = 0.3;
            shape.PictureFormat.Crop.PictureOffsetX = 10;
            shape.PictureFormat.Crop.PictureOffsetY = -5;
            shape.PictureFormat.Crop.PictureWidth = 120;
            shape.PictureFormat.Crop.PictureHeight = 80;
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

        public override string[] UsedResources
        {
            get
            {
                return new string[] { "logo.png" };
            }
        }
    }
}
