using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Formatting.Protection
{
    public class SetRangeLocked : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            //config range B1's Locked property.
            worksheet.Range["B1"].Locked = false;
            //protect worksheet, range B1 can be modified in exported xlsx file.
            worksheet.Protection = true;
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
