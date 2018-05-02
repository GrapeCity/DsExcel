using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting.PdfPageSetup
{
    public class ConfigureHeaderFooter : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet sheet = workbook.Worksheets[0];

            //Set data.
            sheet.Range["A1:G10"].Value = "Text";

            //Set page header.
            sheet.PageSetup.LeftHeader = "&\"Arial,Italic\"LeftHeader";
            sheet.PageSetup.RightHeader = "&KFF0000GrapeCity";
            sheet.PageSetup.CenterHeader = "&P";

            //Set page footer picture.
            System.IO.Stream stream = this.GetResourceStream("logo.png");
            sheet.PageSetup.CenterFooter = "&G";
            sheet.PageSetup.CenterFooterPicture.SetGraphicStream(stream, Drawing.ImageType.PNG);
            sheet.PageSetup.CenterFooterPicture.Width = 100;
            sheet.PageSetup.CenterFooterPicture.Height = 13;
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
