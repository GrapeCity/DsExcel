using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting.PdfPageSetup
{
    public class ConfigureDraft : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet sheet = workbook.Worksheets[0];

            //Set text.
            sheet.Range["A1:G10"].Value = "Text";

            //Add picture in sheet.
            System.IO.Stream stream = this.GetResourceStream("logo.png");
            GrapeCity.Documents.Excel.Drawing.IShape picture = sheet.Shapes.AddPicture(stream, GrapeCity.Documents.Excel.Drawing.ImageType.PNG, 20, 20, 395, 60);

            //Add header graphic.
            System.IO.Stream stream1 = this.GetResourceStream("logo.png");
            sheet.PageSetup.CenterHeader = "&G";
            sheet.PageSetup.CenterHeaderPicture.SetGraphicStream(stream1, Drawing.ImageType.PNG);
            sheet.PageSetup.CenterHeaderPicture.Width = 100;
            sheet.PageSetup.CenterHeaderPicture.Height = 13;

            //Set print without graphics in page content area.
            sheet.PageSetup.Draft = true;
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

        public override string[] UsedResources
        {
            get
            {
                return new string[] { "logo.png" };
            }
        }
    }
}
