using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Comments
{
    public class ClearComment : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            worksheet.Range["C3"].AddComment("Range C3's comment.");
            worksheet.Range["D4"].AddComment("Range D4's comment.");
            worksheet.Range["D5"].AddComment("Range D5's comment.");

            //delete a single cell comment.
            worksheet.Range["D5"].Comment.Delete();

            //clear a range of cells comment.
            worksheet.Range["C3:D4"].ClearComments();
        }
    }
}
