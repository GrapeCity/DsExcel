using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Workbook
{
    public class ImportExcelFileWithPassword : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //Change the path to the real file path when open.
            XlsxOpenOptions options = new XlsxOpenOptions
            {
                Password = "123456"
            };

            workbook.Open(System.IO.Path.Combine(this.CurrentDirectory, "source.xlsx"), options);

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
