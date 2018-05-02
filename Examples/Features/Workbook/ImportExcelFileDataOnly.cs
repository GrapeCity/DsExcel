using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Workbook
{
    public class ImportExcelFileDataOnly : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //Use XlsxOpenOptions.ImportFlags to control what you want to import from excel, ImportFlags.Data means only the data will be imported
            //Change the path to the real file path when open.
            XlsxOpenOptions options = new XlsxOpenOptions();
            options.ImportFlags = ImportFlags.Data;

            workbook.Open(this.CurrentDirectory + "source.xlsx", options);

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

        public override bool IsUpdate
        {
            get
            {
                return true;
            }
        }
    }
}
