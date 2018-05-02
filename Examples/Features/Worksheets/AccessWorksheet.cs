using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Worksheets
{
    public class AccessWorksheet : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //Use sheet index to get worksheet.
            IWorksheet worksheet = workbook.Worksheets[0];

            //Use sheet name to get worksheet.
            IWorksheet worksheet1 = workbook.Worksheets["Sheet1"];
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
