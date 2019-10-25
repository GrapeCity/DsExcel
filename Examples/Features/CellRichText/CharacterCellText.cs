using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.CellRichText
{
    public class CharacterCellText : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            IRange b2 = worksheet.Range["B2"];
            b2.Value = "GrapeCity Documents for Excel";
            b2.Font.Size = 26;
            b2.EntireRow.RowHeight = 42;

            // customize the 'GrapeCity' run
            ITextRun run1 = b2.Characters(0, 9);
            run1.Font.Name = "Agency FB";
            run1.Font.ThemeColor = ThemeColor.Accent1;
            run1.Font.Bold = true;

            // customize the 'Documents' run
            ITextRun run2 = b2.Characters(10, 9);
            run2.Font.ThemeColor = ThemeColor.Accent2;
            run2.Font.Name = "Arial Black";
            run2.Font.Underline = UnderlineType.Single;

            // customize the 'for' run
            ITextRun run3 = b2.Characters(20, 3);
            run3.Font.Italic = true;

            // customize the 'Excel' run
            ITextRun run4 = b2.Characters(24, 5);
            run4.Font.Color = System.Drawing.Color.Blue;
            run4.Font.Bold = true;
        }
    }
}
