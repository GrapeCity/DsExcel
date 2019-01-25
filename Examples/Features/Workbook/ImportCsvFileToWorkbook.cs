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
            Stream stream = this.GetResourceStream("xlsx\\Information.csv");

            //Open csv file stream.
            workbook.Open(stream, OpenFileFormat.Csv);
        }

        public override string TemplateName
        {
            get
            {
                return "Information.csv";
            }
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
        public override string[] UsedResources
        {
            get
            {
                return new string[] { "xlsx\\Information.csv"};
            }
        }
    }

}
