using GrapeCity.Documents.Drawing;
using GrapeCity.Documents.Imaging;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting
{
    public class ConvertSheetToImage : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook, MemoryStream outputStream)
        {
            Stream fileStream = this.GetResourceStream("xlsx\\Employee absence schedule.xlsx");
            workbook.Open(fileStream);

            IWorksheet worksheet = workbook.Worksheets[0];

            //NOTE: To use this feature, you should have valid license for GrapeCity Documents for PDF.
            //Create a pdf document.
            GrapeCity.Documents.Pdf.GcPdfDocument doc = new GrapeCity.Documents.Pdf.GcPdfDocument();
            GrapeCity.Documents.Pdf.Page page = doc.NewPage();

            //Create a PrintManager.
            GrapeCity.Documents.Excel.PrintManager printManager = new GrapeCity.Documents.Excel.PrintManager();

            //Get the first print area of the worksheet.
            IRange printArea = printManager.GetPrintAreas(worksheet)[0];
            //Get the size of the printArea.
            SizeF size = printManager.GetSize(printArea);

            //Modify the size of the page.
            page.Size = size;

            // Draw the printArea to the specified location on the page. 
            printManager.Draw(page, new PointF(0, 0), printArea);

            // Saves the page as an image to a stream.
            page.SaveAsPng(outputStream, new Pdf.SaveAsImageOptions());
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

        public override string TemplateName
        {
            get
            {
                return "Employee absence schedule.xlsx";
            }
        }
        public override string[] UsedResources
        {
            get
            {
                return new string[] { "xlsx\\Employee absence schedule.xlsx" };
            }
        }
    }
}
