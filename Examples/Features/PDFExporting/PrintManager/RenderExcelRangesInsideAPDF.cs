using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting.PrintManager
{
    public class RenderExcelRangesInsideAPDF : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook, MemoryStream outputStream)
        {
            Stream fileStream = this.GetResourceStream("xlsx\\FinancialReport.xlsx");
            workbook.Open(fileStream);
            IWorksheet worksheet = workbook.Worksheets[0];

            //NOTE: To use this feature, you should have valid license for GrapeCity Documents for PDF.
            //Create a pdf document.
            GrapeCity.Documents.Pdf.GcPdfDocument doc = new GrapeCity.Documents.Pdf.GcPdfDocument();
            doc.Load(this.GetResourceStream("Acme-Financial Report 2018.pdf"));

            //Create a PrintManager.
            GrapeCity.Documents.Excel.PrintManager printManager = new GrapeCity.Documents.Excel.PrintManager();

            //Draw the contents of the sheet3 to the fourth page. 
            IRange printArea1 = workbook.Worksheets[2].Range["A3:C24"];
            SizeF size1 = printManager.GetSize(printArea1);
            RectangleF position1 = doc.FindText(new GrapeCity.Documents.Pdf.FindTextParams("Proposition enhancements are", true, true), new GrapeCity.Documents.Common.OutputRange(4, 4))[0].Bounds.ToRect();
            printManager.Draw(doc.Pages[3], new RectangleF(position1.X + position1.Width + 70, position1.Y, size1.Width, size1.Height), printArea1);

            //Draw the contents of the sheet1 to the fifth page. 
            IRange printArea2 = workbook.Worksheets[0].Range["A4:E29"];
            SizeF size2 = printManager.GetSize(printArea2);
            RectangleF position2 = doc.FindText(new GrapeCity.Documents.Pdf.FindTextParams("expenditure, an improvement in working", true, true), new GrapeCity.Documents.Common.OutputRange(5, 5))[0].Bounds.ToRect();
            printManager.Draw(doc.Pages[4], new RectangleF(position2.X, position2.Y + position2.Height + 20, size2.Width, size2.Height), printArea2);

            //Draw the contents of the sheet2 to the sixth page. 
            IRange printArea3 = workbook.Worksheets[1].Range["A2:E28"];
            SizeF size3 = printManager.GetSize(printArea3);
            RectangleF position3 = doc.FindText(new GrapeCity.Documents.Pdf.FindTextParams("company will be able to continue", true, true), new GrapeCity.Documents.Common.OutputRange(6, 6))[0].Bounds.ToRect();
            printManager.Draw(doc.Pages[5], new RectangleF(position3.X, position3.Y + position3.Height + 20, doc.Pages[5].Size.Width - position3.X * 2 - 10, size3.Height), printArea3);

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

        public override string TemplateName
        {
            get
            {
                return "FinancialReport.xlsx";
            }
        }
        public override string[] UsedResources
        {
            get
            {
                return new string[] { "xlsx\\FinancialReport.xlsx", "Acme-Financial Report 2018.pdf" };
            }
        }
    }
}
