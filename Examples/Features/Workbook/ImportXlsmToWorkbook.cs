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
            workbook.Open(this.CurrentDirectory + "macros.xlsm");

            // Macros can be preserved after saving
            workbook.Save(this.CurrentDirectory + "macros-exported.xlsm");

        }

        public override bool IsNew => true;

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
