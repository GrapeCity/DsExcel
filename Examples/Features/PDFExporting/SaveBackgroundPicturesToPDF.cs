using GrapeCity.Documents.Excel.Drawing;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting
{
    public class SaveBackgroundPicturesToPDF : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            var fileStream = this.GetResourceStream("xlsx\\To_Do_List.xlsx");
            workbook.Open(fileStream);

            IWorksheet worksheet = workbook.Worksheets[0];

            Stream stream = this.GetResourceStream("AcmeLogo.png");

            //Add a background picture for the worksheet, and the background picture will be rendered into the destination rectangle[10, 10, 500, 370].
            IBackgroundPicture picture = worksheet.BackgroundPictures.AddPictureInPixel(stream, ImageType.PNG, 10, 10, 500, 370);

            //The background picture will be resized to fill the destination dimensions.
            picture.BackgroundImageLayout = ImageLayout.Tile;

            //Sets the transparency of the background picture.
            picture.Transparency = 0.50;
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

        public override string[] UsedResources
        {
            get
            {
                return new string[] { "xlsx\\To_Do_List.xlsx", "AcmeLogo.png" };
            }
        }
    }
}
