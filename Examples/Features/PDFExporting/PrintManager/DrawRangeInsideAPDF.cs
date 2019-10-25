using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting.PrintManager
{
    public class DrawRangeInsideAPDF : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook, MemoryStream outputStream)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            //Set value
            worksheet.Range["A4:C4"].Value = new string[] { "Device", "Quantity", "Unit Price" };
            worksheet.Range["A5:C8"].Value = new object[,] {
                { "T540p", 12, 9850 },
                { "T570", 5, 7460 },
                { "Y460", 6, 5400 },
                { "Y460F", 8, 6240 }
            };

            //Set style
            worksheet.Range["A4:C4"].Font.Bold = true;
            worksheet.Range["A4:C4"].Font.Color = Color.White;
            worksheet.Range["A4:C4"].Interior.Color = Color.LightBlue;
            worksheet.Range["A5:C8"].Borders[BordersIndex.InsideHorizontal].Color = Color.Orange;
            worksheet.Range["A5:C8"].Borders[BordersIndex.InsideHorizontal].LineStyle = BorderLineStyle.DashDot;

            //NOTE: To use this feature, you should have valid license for GrapeCity Documents for PDF.
            //Create a pdf document.
            GrapeCity.Documents.Pdf.GcPdfDocument doc = new GrapeCity.Documents.Pdf.GcPdfDocument();
            GrapeCity.Documents.Pdf.Page page = doc.NewPage();
            GrapeCity.Documents.Pdf.GcPdfGraphics g = page.Graphics;

            //Create a PrintManager.
            GrapeCity.Documents.Excel.PrintManager printManager = new GrapeCity.Documents.Excel.PrintManager();

            // Draw the Range"A4:C8" to the specified location on the page. 
            printManager.Draw(page, new PointF(30, 100), worksheet.Range["A4:C8"]);

            //Save the modified pages into pdf file.
            doc.Save(outputStream);
        }

        public override bool SavePageInfos

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
