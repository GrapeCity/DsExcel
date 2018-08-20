using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting.PdfPageSetup
{
    public class ConfigurePrintPageRange : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet sheet = workbook.Worksheets[0];

            //Set pages' data.
            sheet.Range["A1:J46"].Value = "Page1";
            sheet.Range["A1:J46"].Interior.Color = Color.LightGreen;

            sheet.Range["A47:J92"].Value = "Page2";
            sheet.Range["A47:J92"].Interior.Color = Color.LightYellow;

            sheet.Range["K1:T46"].Value = "Page3";
            sheet.Range["K1:T46"].Interior.Color = Color.OrangeRed;

            sheet.Range["K47:T92"].Value = "Page4";
            sheet.Range["K47:T92"].Interior.Color = Color.DarkOrange;

            sheet.Range["U1:AD46"].Value = "Page5";
            sheet.Range["U1:AD46"].Interior.Color = Color.LightBlue;

            sheet.Range["U47:AD92"].Value = "Page6";
            sheet.Range["U47:AD92"].Interior.Color = Color.IndianRed;
            sheet.PageSetup.PrintHeadings = true;

            //Set print page range, print p1, p3 to p5.
            sheet.PageSetup.PrintPageRange = "1,3-5";
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
