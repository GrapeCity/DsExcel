using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.CellRichText
{
    public class ConfigRunFont : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            IRange a2 = worksheet.Range["A2"];
            a2.Font.Size = 18;
            a2.Font.Bold = true;
            a2.VerticalAlignment = VerticalAlignment.Center;

            a2.EntireRow.RowHeight = 42;
            a2.EntireColumn.ColumnWidth = 40;
            a2.Value = "Perfect square trinomial";

            ITextRun run1 = a2.Characters(8, 7);
            run1.Font.Italic = true;
            run1.Font.ThemeColor = ThemeColor.Accent1;

            IRange b2 = worksheet.Range["B2"];
            b2.Font.Size = 26;
            b2.EntireColumn.ColumnWidth = 60;

            b2.Value = "(a+b)2 = a2+2ab+b2";
            
            ITextRun superRun1 = b2.Characters(5, 1);
            superRun1.Font.Superscript = true;
            superRun1.Font.Color = System.Drawing.Color.Red;

            ITextRun superRun2 = b2.Characters(10, 1);
            superRun2.Font.Superscript = true;
            superRun2.Font.Color = System.Drawing.Color.Green;

            ITextRun superRun3 = b2.Characters(17, 1);
            superRun3.Font.Superscript = true;
            superRun3.Font.Color = System.Drawing.Color.Blue;
           
        }
    }
}
