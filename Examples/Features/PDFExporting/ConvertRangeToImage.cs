using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting
{
    public class ConvertRangeToImage : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook, MemoryStream outputStream)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            //Set value
            worksheet.Range["A1:C1"].Value = new string[] { "Device", "Quantity", "Unit Price" };
            worksheet.Range["A2:C5"].Value = new object[,] {
                { "T540p", 12, 9850 },
                { "T570", 5, 7460 },
                { "Y460", 6, 5400 },
                { "Y460F", 8, 6240 }
            };

            //Set style
            worksheet.Range["A1:C1"].Font.Bold = true;
            worksheet.Range["A1:C1"].Font.Color = Color.White;
            worksheet.Range["A1:C1"].Interior.Color = Color.LightBlue;
            worksheet.Range["A2:C5"].Borders[BordersIndex.InsideHorizontal].Color = Color.Orange;
            worksheet.Range["A2:C5"].Borders[BordersIndex.InsideHorizontal].LineStyle = BorderLineStyle.DashDot;

            //NOTE: To use this feature, you should have valid license for GrapeCity Documents for PDF.
            //Create a pdf document.
            GrapeCity.Documents.Pdf.GcPdfDocument doc = new GrapeCity.Documents.Pdf.GcPdfDocument();
            GrapeCity.Documents.Pdf.Page page = doc.NewPage();

            //Create a PrintManager.
            GrapeCity.Documents.Excel.PrintManager printManager = new GrapeCity.Documents.Excel.PrintManager();

            //Get the size of the range"A1:C5".
            SizeF size = printManager.GetSize(worksheet.Range["A1:C5"]);

            //Modify the size of the page.
            page.Size = size;

            // Draw the Range"A1:E5" to the specified location on the page. 
            printManager.Draw(page, new PointF(0, 0), worksheet.Range["A1:C5"]);

            // Saves the page as an image to a stream.
            page.SaveAsPng(outputStream, new Pdf.SaveAsImageOptions() { Resolution = 72 });
        }

        public override bool SaveAsImages
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
