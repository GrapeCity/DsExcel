using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting
{
    public class SavePictureToPDF : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            worksheet.PageSetup.Orientation = PageOrientation.Landscape;
            System.IO.Stream stream = this.GetResourceStream("logo.png");
            GrapeCity.Documents.Excel.Drawing.IShape picture = worksheet.Shapes.AddPicture(stream, GrapeCity.Documents.Excel.Drawing.ImageType.PNG, 20, 20, 690, 100);

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

        public override string[] UsedResources => new string[] { "logo.png" };
    }
}
