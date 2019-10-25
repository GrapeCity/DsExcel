using GrapeCity.Documents.Excel.Drawing;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting.ExportShape
{
    public class ShapeWithText : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            // Add a rectangle
            IShape rectangle = worksheet.Shapes.AddShape(AutoShapeType.Rectangle, 50, 30, 300, 200);
            
            // Add rich text to rectangle
            rectangle.Fill.Color.RGB = System.Drawing.Color.White;

            // Add first paragraph
            ITextRange run1 = rectangle.TextFrame.TextRange.Paragraphs[0].Runs.Add("         Doc");
            run1.Font.Color.RGB = System.Drawing.Color.Tomato;
            ITextRange run2 = rectangle.TextFrame.TextRange.Paragraphs[0].Runs.Add("ume");
            run2.Font.Color.RGB = System.Drawing.Color.Yellow;
            ITextRange run3 = rectangle.TextFrame.TextRange.Paragraphs[0].Runs.Add("nts");
            run3.Font.Color.RGB = System.Drawing.Color.LightBlue;
            ITextRange paragraph0 = rectangle.TextFrame.TextRange.Paragraphs[0];
            paragraph0.Font.Size = 36;
            paragraph0.Font.Bold = true;

            rectangle.TextFrame.TextRange.Paragraphs.Add(" ");

            // Add second paragraph
            ITextRange paragraph1 = rectangle.TextFrame.TextRange.Paragraphs.Add();
            ITextRange run4 = paragraph1.Runs.Add("Fast, efficient Excel, Word, and PDF APIs for .NET Standard 2.0 and Java");
            run4.Font.Color.RGB = System.Drawing.Color.Black;
            run4.Font.Size = 20;
            run4.Font.Name = "Agency FB";

            rectangle.TextFrame.TextRange.Paragraphs.Add(" ");

            // Add third paragraph
            ITextRange paragraph2 = rectangle.TextFrame.TextRange.Paragraphs.Add();
            ITextRange run5 = paragraph2.Runs.Add("Take total document control with ultra-fast, low-footprint document APIs for enterprise apps");
            run5.Font.Color.RGB = System.Drawing.Color.Black;
            run5.Font.Size = 16;
            run5.Font.Name = "Times New Roman";
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
