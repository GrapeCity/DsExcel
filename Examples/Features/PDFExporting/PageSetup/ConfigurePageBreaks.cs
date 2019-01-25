using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting.PdfPageSetup
{
    public class ConfigurePageBreaks : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet sheet = workbook.Worksheets[0];

            sheet.Range["A1:B5"].Value = new object[,]
            {
                {1, 2},
                {3, 4},
                {5, 6},
                {7, 8},
                {9, 10}
            };

            //Add page break
            sheet.HPageBreaks.Add(sheet.Range["B3"]);
            sheet.VPageBreaks.Add(sheet.Range["B3"]);
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
    }
}
