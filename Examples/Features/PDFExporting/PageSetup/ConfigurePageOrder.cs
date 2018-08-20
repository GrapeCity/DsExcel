using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting.PdfPageSetup
{
    public class ConfigurePageOrder : ExampleBase
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

            sheet.PageSetup.PrintHeadings = true;
            
            //Set page order. Now the page order is p1-p3-p2-p4. Origin is p1-p2-p3-p4.
            sheet.PageSetup.Order = Order.OverThenDown;
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
