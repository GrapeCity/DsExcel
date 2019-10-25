using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Workbook
{
    public class ImportXlsmToWorkbook : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            // GcExcel supports open xlsm file
            workbook.Open(System.IO.Path.Combine(this.CurrentDirectory, "macros.xlsm"));

            // Macros can be preserved after saving
            workbook.Save(System.IO.Path.Combine(this.CurrentDirectory, "macros-exported.xlsm"));

        }

        public override bool CanDownload
        {
            get
            {
                return false;
            }
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
