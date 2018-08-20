using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Comments
{
    public class ConfigCommentLayout : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            IComment commentC3 = worksheet.Range["C3"].AddComment("Range C3's comment.");
            commentC3.Shape.Line.Color.RGB = Color.LightGreen;
            commentC3.Shape.Line.Weight = 3;
            commentC3.Shape.Line.Style = GrapeCity.Documents.Excel.Drawing.LineStyle.ThickThin;
            commentC3.Shape.Line.DashStyle = GrapeCity.Documents.Excel.Drawing.LineDashStyle.Solid;
            commentC3.Shape.Fill.Color.RGB = Color.Pink;
            commentC3.Shape.Width = 100;
            commentC3.Shape.Height = 200;
            commentC3.Shape.TextFrame.TextRange.Font.Bold = true;
            commentC3.Visible = true;
        }
    }
}
