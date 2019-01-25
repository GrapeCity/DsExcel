using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Workbook
{
    public class ImportExcelFileWithoutCalculation : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //When XlsxOpenOptions.DoNotRecalculateAfterOpened means GrapeCity Documents for Excel will just read all the cached values without calculating again after
            //opening an Excel file.
            //Change the path to the real file path when open.

            XlsxOpenOptions options = new XlsxOpenOptions();
            options.DoNotRecalculateAfterOpened = true;

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
    }
}
