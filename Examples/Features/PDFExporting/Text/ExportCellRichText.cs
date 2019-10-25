using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting.Text
{
    public class ExportCellRichText : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            IRange a1 = worksheet.Range["A1"];
            a1.Value = "Perfect square trinomial";
            a1.Font.Size = 26;
            a1.Font.Bold = true;
            a1.VerticalAlignment = VerticalAlignment.Bottom;

            a1.EntireRow.RowHeight = 42;
            a1.EntireColumn.ColumnWidth = 50;

            ITextRun run1 = a1.Characters(8, 7);
            run1.Font.Italic = true;
            run1.Font.ThemeColor = ThemeColor.Accent1;

            IRange b1 = worksheet.Range["B1"];
            b1.Font.Size = 22;
            b1.EntireColumn.ColumnWidth = 40;

            b1.Value = "(a+b)2 = a2+2ab+b2";
            b1.VerticalAlignment = VerticalAlignment.Center;
            
            ITextRun superRun1 = b1.Characters(5, 1);
            superRun1.Font.Superscript = true;
            superRun1.Font.Color = System.Drawing.Color.Red;

            ITextRun superRun2 = b1.Characters(10, 1);
            superRun2.Font.Superscript = true;
            superRun2.Font.Color = System.Drawing.Color.Green;

            ITextRun superRun3 = b1.Characters(17, 1);
            superRun3.Font.Superscript = true;
            superRun3.Font.Color = System.Drawing.Color.Blue;
           
        }

        public override bool SavePdf => true;

        public override bool ShowViewer => false;
    }
}
