using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Comments
{
    public class AddComment : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            //create comment for range C3.
            IComment comment = worksheet.Range["C3"].AddComment("Range C3's comment.");

            //change comment's text.
            comment.Text = "Range C3's new comment.";
        }
    }
}
