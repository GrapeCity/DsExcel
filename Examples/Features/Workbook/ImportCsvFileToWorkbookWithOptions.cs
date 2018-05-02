using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Workbook
{
    public class ImportCsvFileToWorkbookWithOptions : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //Open csv with more settings.
            CsvOpenOptions options = new CsvOpenOptions();
            options.SeparatorString = "-";

            //Change the path to the real file path when open.
            workbook.Open(this.CurrentDirectory + "source.csv", options);

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

        public override bool IsNew
        {
            get
            {
                return true;
            }
        }
    }

}
