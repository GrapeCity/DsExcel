using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Shape
{
    public class ConfigShapeWithPatternFill : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            GrapeCity.Documents.Excel.Drawing.IShape shape = worksheet.Shapes.AddShape(GrapeCity.Documents.Excel.Drawing.AutoShapeType.Parallelogram, 1, 1, 200, 100);
            shape.Fill.Patterned(GrapeCity.Documents.Excel.Drawing.PatternType.Percent10);
            shape.Fill.Color.ObjectThemeColor = ThemeColor.Accent2;
            shape.Fill.PatternColor.ObjectThemeColor = ThemeColor.Accent6;
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
