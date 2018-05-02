using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting.Text
{
    public class Overflow : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet sheet = workbook.Worksheets[0];

            sheet.Range["F2, F4"].Value = "This is a test string of overflow";

            sheet.Range["F6, F8"].Value = "This is a test string of overflow with right alignment";
            sheet.Range["F6, F8"].HorizontalAlignment = HorizontalAlignment.Right;

            sheet.Range["D8, H4"].Value = 123;

            //Other settings
            sheet.Range["A1:J10"].Borders.LineStyle = BorderLineStyle.Thin;
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

        public override bool IsNew
        {
            get
            {
                return true;
            }
        }
    }
}
