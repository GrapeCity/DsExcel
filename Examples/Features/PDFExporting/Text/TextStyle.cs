using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting.Text
{
    public class TextStyle : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet sheet = workbook.Worksheets[0];

            #region Aligment
            sheet.Range["A1"].Value = "Alignment";

            sheet.Range["B2"].Value = "Left Alignment";
            sheet.Range["B2"].HorizontalAlignment = HorizontalAlignment.Left;

            sheet.Range["C2"].Value = "Center Alignment";
            sheet.Range["C2"].HorizontalAlignment = HorizontalAlignment.Center;

            sheet.Range["D2"].Value = "Right Alignment";
            sheet.Range["D2"].HorizontalAlignment = HorizontalAlignment.Right;

            sheet.Range["B3"].Value = "Top Alignment";
            sheet.Range["B3"].VerticalAlignment = VerticalAlignment.Top;

            sheet.Range["C3"].Value = "Middle Alignment";
            sheet.Range["C3"].VerticalAlignment = VerticalAlignment.Center;

            sheet.Range["D3"].Value = "Bottom Alignment";
            sheet.Range["D3"].VerticalAlignment = VerticalAlignment.Bottom;

            sheet.Range["B4"].Value = "This is  a test string for Justify Alignment. \nThis is a test string for Justify Alignment. ";
            sheet.Range["B4"].HorizontalAlignment = HorizontalAlignment.Justify;
            sheet.Range["B4"].VerticalAlignment = VerticalAlignment.Justify;

            sheet.Range["C4"].Value = "This is  a test string for Distributed Alignment. \nThis is a test string for Distributed Alignment. ";
            sheet.Range["C4"].HorizontalAlignment = HorizontalAlignment.Distributed;
            sheet.Range["C4"].VerticalAlignment = VerticalAlignment.Distributed;
            #endregion

            //Wordwrap
            sheet.Range["A6"].Value = "Wordwrap";
            sheet.Range["B7"].Value = "This is a test string for Wordwrap";
            sheet.Range["C7"].Value = "This is a test string \n for Wordwrap";
            sheet.Range["B7:C7"].WrapText = true;

            //Indent
            sheet.Range["A9"].Value = "Indent";
            sheet.Range["B10"].Value = "Left Indent";
            sheet.Range["B10"].IndentLevel = 3;
            sheet.Range["C10"].Value = "Right Indent";
            sheet.Range["C10"].IndentLevel = 3;
            sheet.Range["C10"].HorizontalAlignment = HorizontalAlignment.Right;

            //Shrink to fit
            sheet.Range["A12"].Value = "Shrink to fit";
            sheet.Range["B13"].Value = "This is a test string for \"Shrink to fit\"";
            sheet.Range["B13"].ShrinkToFit = true;

            //Underline
            sheet.Range["A15"].Value = "Underline";
            sheet.Range["B16"].Value = "Single Underline";
            sheet.Range["B16"].Font.Underline = UnderlineType.Single;

            //Strikthrough
            sheet.Range["A18"].Value = "Strikthrough";
            sheet.Range["B19"].Value = "Strikthrough";
            sheet.Range["B19"].Font.Strikethrough = true;

            //Other settings
            sheet.Columns[0].Font.Bold = true;
            sheet.Columns[0].ColumnWidthInPixel = 100;
            sheet.Columns[1].ColumnWidthInPixel = 200;
            sheet.Columns[2].ColumnWidthInPixel = 245;
            sheet.Columns[3].ColumnWidthInPixel = 234;
            sheet.Rows[2].RowHeightInPixel = 72;
            sheet.Rows[3].RowHeightInPixel = 123;
            sheet.Rows[6].RowHeightInPixel = 48;

            sheet.Range["A1:D19"].Borders.LineStyle = BorderLineStyle.Thin;
            sheet.PageSetup.PaperSize = PaperSize.A3;
        }

        public override bool SavePdf
        {
            get
            {
                return true;
            }
        }

        public override bool ShowViewer
        {
            get
            {
                return false;
            }
        }
    }
}
