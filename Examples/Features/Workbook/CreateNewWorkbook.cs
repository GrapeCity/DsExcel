using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Workbook
{
    public class CreateNewWorkbook : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //Create empty workbook, contains one worksheet default.
            GrapeCity.Documents.Excel.Workbook workbookNew = new GrapeCity.Documents.Excel.Workbook();
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
