using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Workbook
{
    public class UnprotectWorkbook : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Range["A1"].Value = "GrapeCity Documents for Excel";

            //Protects the workbook so that other users cannot view hidden worksheets, add, move, delete, hidie, or rename worksheets.
            //The protection only happens when you open it with an Excel application.
            workbook.Protect();

            //Removes the above protection from the workbook.
            workbook.Unprotect();
        }

        public override bool CanDownload
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

        public override bool IsNew
        {
            get
            {
                return true;
            }
        }
    }
}
