using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Formatting
{
    public class ApplyStyle : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            //Change to build in name style.
            worksheet.Range["A1"].Value = "Bad";
            worksheet.Range["A1"].Style = workbook.Styles["Bad"];


            //Change to custom name style.
            //Add custom name style.
            IStyle style = workbook.Styles.Add("testStyle");

            //Config custom name style settings begin.
            //Border
            style.Borders[BordersIndex.EdgeLeft].LineStyle = BorderLineStyle.Thin;
            style.Borders[BordersIndex.EdgeTop].LineStyle = BorderLineStyle.Thick;
            style.Borders[BordersIndex.EdgeRight].LineStyle = BorderLineStyle.Double;
            style.Borders[BordersIndex.EdgeBottom].LineStyle = BorderLineStyle.Double;
            style.Borders.Color = Color.FromRGB(0, 255, 0);

            //Font
            style.Font.ThemeColor = ThemeColor.Accent1;
            style.Font.TintAndShade = 0.8;
            style.Font.Italic = true;
            style.Font.Bold = true;
            style.Font.Name = "LiSu";
            style.Font.Size = 28;
            style.Font.Strikethrough = true;
            style.Font.Subscript = true;
            style.Font.Superscript = false;
            style.Font.Underline = UnderlineType.Double;

            //Protection
            style.FormulaHidden = true;
            style.Locked = false;

            //Number
            style.NumberFormat = "#,##0_);[Red](#,##0)";

            //Alignment
            style.HorizontalAlignment = HorizontalAlignment.Right;
            style.VerticalAlignment = VerticalAlignment.Bottom;
            style.WrapText = true;
            style.IndentLevel = 5;
            style.Orientation = 45;

            //Fill
            style.Interior.ColorIndex = 5;
            style.Interior.Pattern = GrapeCity.Documents.Excel.Pattern.Down;
            style.Interior.PatternColor = Color.FromRGB(0, 0, 255);


            style.IncludeAlignment = false;
            style.IncludeBorder = true;
            style.IncludeFont = false;
            style.IncludeNumber = true;
            style.IncludePatterns = false;
            style.IncludeProtection = true;
            //Config custom name style settings end.

            //Set range's style to custom name style.
            worksheet.Range["A2"].Value = "My test style";
            worksheet.Range["A2"].Style = worksheet.Workbook.Styles["testStyle"];

        }
    }
}
