using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Shape
{
    public class ConfigShapeWithTextureFill : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            GrapeCity.Documents.Excel.Drawing.IShape shape = worksheet.Shapes.AddShape(GrapeCity.Documents.Excel.Drawing.AutoShapeType.Heart, 1, 1, 100, 100);
            shape.Fill.PresetTextured(GrapeCity.Documents.Excel.Drawing.PresetTexture.Canvas);
            shape.Fill.TextureAlignment = GrapeCity.Documents.Excel.Drawing.TextureAlignment.Center;
            shape.Fill.TextureOffsetX = 2.5;
            shape.Fill.TextureOffsetY = 3.2;
            shape.Fill.TextureHorizontalScale = 0.9;
            shape.Fill.TextureVerticalScale = 0.2;
            shape.Fill.Transparency = 0.5;
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
