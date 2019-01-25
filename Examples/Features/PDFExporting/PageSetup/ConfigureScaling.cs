using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting.PdfPageSetup
{
    public class ConfigureScaling : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet sheet = workbook.Worksheets[0];

            System.IO.Stream stream = this.GetResourceStream("logo.png");
            GrapeCity.Documents.Excel.Drawing.IShape picture = sheet.Shapes.AddPicture(stream, GrapeCity.Documents.Excel.Drawing.ImageType.PNG, 20, 20, 395, 60);
            sheet.Range["B2:D4"].Value = "Text";

            sheet.PageSetup.PrintGridlines = true;

            //Set scaling.
            sheet.PageSetup.Zoom = 200;
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
