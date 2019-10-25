using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting.PdfPageSetup
{
    public class ConfigurePrintTails : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            var fileStream = this.GetResourceStream("xlsx\\RepeatTails.xlsx");
            workbook.Open(fileStream);

            //Repeat the columns "I" at the left of each page when saving worksheet2 to pdf.
            IWorksheet worksheet1 = workbook.Worksheets[0];
            worksheet1.PageSetup.PrintTailColumns = "$I:$I";

            //Repeat the row 50th at the bottom of each page when saving worksheet1 to pdf.
            IWorksheet worksheet2 = workbook.Worksheets[1];
            worksheet2.PageSetup.PrintTailRows = "$50:$50";
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
                return "RepeatTails.xlsx";
            }
        }
        public override string[] UsedResources
        {
            get
            {
                return new string[] { "xlsx\\RepeatTails.xlsx" };
            }
        }
    }
}
