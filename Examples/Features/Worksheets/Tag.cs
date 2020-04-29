using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Worksheets
{
    public class Tag : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            //Add tag for worksheet
            worksheet.Tag = "This is a Tag for sheet.";

            //Add tag for range A1:B2
            worksheet.Range["A1:B2"].Tag = "This is a Tag for A1:B2";

            //Add tag for row 4
            worksheet.Range["A4"].EntireRow.Tag = "This is a Tag for Row 4";

            //Add tag for column F
            worksheet.Range["F5"].EntireColumn.Tag = "This is a Tag for Column F";
        }

        public override bool ShowViewer
        {
            get
            {
                return false;
            }
        }

        public override bool CanDownload
        {
            get
            {
                return false;
            }
        }
    }
}
