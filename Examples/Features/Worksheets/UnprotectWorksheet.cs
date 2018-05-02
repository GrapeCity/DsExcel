using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Worksheets
{
    public class UnprotectWorksheet : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            //protect worksheet, allow insert column.
            worksheet.Protection = true;
            worksheet.ProtectionSettings.AllowInsertingColumns = true;

            //Unprotect worksheet.
            worksheet.Protection = false;
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
