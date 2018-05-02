using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Shape
{
    public class ConfigShapeLine : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            GrapeCity.Documents.Excel.Drawing.IShape shape = worksheet.Shapes.AddShape(GrapeCity.Documents.Excel.Drawing.AutoShapeType.Heart, 1, 1, 100, 100);
            shape.Line.DashStyle = GrapeCity.Documents.Excel.Drawing.LineDashStyle.Dash;
            shape.Line.Style = GrapeCity.Documents.Excel.Drawing.LineStyle.Single;
            shape.Line.Weight = 2;
            shape.Line.Color.ObjectThemeColor = ThemeColor.Accent6;
            shape.Line.Transparency = 0.3;
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
