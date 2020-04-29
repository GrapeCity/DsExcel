using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Worksheets
{
    public class UnprotectWorksheetWithPassword : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            var fileStream = this.GetResourceStream("xlsx\\Medical office start-up expenses 1.xlsx");
            workbook.Open(fileStream);

            //Use a password to protect all the worksheet. If you forget the password, you cannot unprotect the worksheet.
            foreach (IWorksheet worksheet in workbook.Worksheets)
            {
                worksheet.Protect("Y6dh!et5");
            }

            //Use the correct password to remove the above protection from the worksheet.
            foreach (IWorksheet worksheet in workbook.Worksheets)
            {
                worksheet.Unprotect("Y6dh!et5");
            }
        }

        public override bool ShowViewer
        {
            get
            {
                return false;
            }
        }

        public override string[] UsedResources
        {
            get
            {
                return new string[] { "xlsx\\Medical office start-up expenses 1.xlsx" };
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
