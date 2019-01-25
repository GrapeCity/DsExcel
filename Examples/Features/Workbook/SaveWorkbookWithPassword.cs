using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Workbook
{
    public class SaveWorkbookWithPassword : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //Change the path to real export path when save.
            XlsxSaveOptions options = new XlsxSaveOptions();
            options.Password = "123456";

            workbook.Save(this.CurrentDirectory + "dest.xlsx", options);

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
