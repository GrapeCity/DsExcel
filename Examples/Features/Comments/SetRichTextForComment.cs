using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Comments
{
    public class SetRichTextForComment : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            IComment commentC3 = worksheet.Range["C3"].AddComment("This is a rich text comment:\r\n");

            //config the paragraph's style.
            commentC3.Shape.TextFrame.TextRange.Paragraphs[0].Font.Bold = true;

            //add runs for the paragraph.
            commentC3.Shape.TextFrame.TextRange.Paragraphs[0].Runs.Add("Run1 font size is 15.", 1);
            commentC3.Shape.TextFrame.TextRange.Paragraphs[0].Runs.Add("Run2 font strikethrough.", 2);
            commentC3.Shape.TextFrame.TextRange.Paragraphs[0].Runs.Add("Run3 font italic, green color.");

            //config the first run of the paragraph's style.
            commentC3.Shape.TextFrame.TextRange.Paragraphs[0].Runs[1].Font.Size = 15;
            //config the second run of the paragraph's style. 
            commentC3.Shape.TextFrame.TextRange.Paragraphs[0].Runs[2].Font.Strikethrough = true;

            //config the third run of the paragraph's style. 
            commentC3.Shape.TextFrame.TextRange.Paragraphs[0].Runs[3].Font.Italic = true;
            commentC3.Shape.TextFrame.TextRange.Paragraphs[0].Runs[3].Font.Color.RGB = Color.Green;

            //show comment.
            commentC3.Visible = true;

            commentC3.Shape.WidthInPixel = 300;
            commentC3.Shape.HeightInPixel = 100;
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
