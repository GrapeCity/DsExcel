using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Shape
{
    public class ConfigShapeWithGradientFill : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            GrapeCity.Documents.Excel.Drawing.IShape shape = worksheet.Shapes.AddShape(GrapeCity.Documents.Excel.Drawing.AutoShapeType.Heart, 1, 1, 100, 100);
            shape.Fill.PresetGradient(GrapeCity.Documents.Excel.Drawing.GradientStyle.Vertical, 3, GrapeCity.Documents.Excel.Drawing.PresetGradientType.Silver);
            shape.Fill.RotateWithObject = false;
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
