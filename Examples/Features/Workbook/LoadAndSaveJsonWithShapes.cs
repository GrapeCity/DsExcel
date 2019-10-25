using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Workbook
{
    public class LoadAndSaveJsonWithShapes : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            GrapeCity.Documents.Excel.Workbook workbookWithShape = new GrapeCity.Documents.Excel.Workbook();

            IWorksheet worksheet = workbookWithShape.Worksheets[0];

            // Add a shape in worksheet
            GrapeCity.Documents.Excel.Drawing.IShape shape = worksheet.Shapes.AddShape(GrapeCity.Documents.Excel.Drawing.AutoShapeType.Parallelogram, 1, 1, 200, 100);
            shape.Line.DashStyle = GrapeCity.Documents.Excel.Drawing.LineDashStyle.Dash;
            shape.Line.Style = GrapeCity.Documents.Excel.Drawing.LineStyle.Single;
            shape.Line.Weight = 2;
            shape.Line.Color.ObjectThemeColor = ThemeColor.Accent6;
            shape.Line.Transparency = 0.3;

            // jsonString contains shapes
            String jsonString = workbookWithShape.ToJson();

            // GcExcel can load json string contains shapes now
            workbook.FromJson(jsonString);
        }

        public override bool CanDownload
        {
            get
            {
                return true;
            }
        }
    }

}
