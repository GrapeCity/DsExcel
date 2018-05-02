using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Workbook
{
    public class ImportCsvFileToWorkbook : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            Stream stream = this.GetTemplateStream("Information.csv");

            //Open csv file stream.
            workbook.Open(stream, OpenFileFormat.Csv);
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
                return true;
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
