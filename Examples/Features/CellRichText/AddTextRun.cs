using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.CellRichText
{
    public class AddTextRun : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            IRange b2 = worksheet.Range["B2"]; 

            // customize the 'GrapeCity' run
            ITextRun run1 = b2.RichText.Add("GrapeCity");
            run1.Font.Name = "Agency FB";
            run1.Font.Size = 26;
            run1.Font.ThemeColor = ThemeColor.Accent1;
            run1.Font.Bold = true;

            // customize the 'Documents' run
            ITextRun run2 = b2.RichText.Add(" Documents");
            run2.Font.ThemeColor = ThemeColor.Accent2;
            run2.Font.Name = "Arial Black";
            run2.Font.Size = 20;

            // customize the 'for' run
            ITextRun run3 = b2.RichText.Add(" for ");
            run3.Font.Italic = true;

            // customize the 'Excel' run
            ITextRun run4 = b2.RichText.Add("Excel");
            run4.Font.Color = System.Drawing.Color.Blue;
            run4.Font.Bold = true;
            run4.Font.Size = 26;
            run4.Font.Underline = UnderlineType.Double;

            b2.EntireRow.RowHeight = 42;
        }
    }

}
