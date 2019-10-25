using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting
{
    public class ConvertExcelToPDF : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            Stream fileStream = this.GetResourceStream("xlsx\\Employee absence schedule.xlsx");
            workbook.Open(fileStream);
        }

        public override bool SavePdf
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
                return false;
            }
        }
        public override string TemplateName
        {
            get
            {
                return "Employee absence schedule.xlsx";
            }
        }
        public override string[] UsedResources
        {
            get
            {
                return new string[] { "xlsx\\Employee absence schedule.xlsx" };
            }
        }
    }
}
