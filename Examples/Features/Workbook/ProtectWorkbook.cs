using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Workbook
{
    public class ProtectWorkbook : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            var fileStream = this.GetResourceStream("xlsx\\Medical office start-up expenses 1.xlsx");
            workbook.Open(fileStream);

            //Protects the workbook with a password so that other users cannot view hidden worksheets, add, move, delete, hide, or rename worksheets.
            //The protection only happens when you open it with an Excel application.
            workbook.Protect("Y6dh!et5");
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
