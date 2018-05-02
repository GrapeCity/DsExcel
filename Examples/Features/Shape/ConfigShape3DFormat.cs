using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Shape
{
    public class ConfigShape3DFormat : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            GrapeCity.Documents.Excel.Drawing.IShape shape = worksheet.Shapes.AddShape(GrapeCity.Documents.Excel.Drawing.AutoShapeType.Parallelogram, 50, 30, 200, 100);
            shape.ThreeD.RotationX = 50;
            shape.ThreeD.RotationY = 20;
            shape.ThreeD.RotationZ = 30;
            shape.ThreeD.Depth = 7;
            shape.ThreeD.Z = 20;
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
