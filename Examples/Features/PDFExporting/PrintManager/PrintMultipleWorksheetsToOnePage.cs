using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting.PrintManager
{
    public class PrintMultipleWorksheetsToOnePage : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook, MemoryStream outputStream)
        {
            Stream fileStream = this.GetResourceStream("xlsx\\Multiple sheets one page.xlsx");
            workbook.Open(fileStream);

            //NOTE: To use this feature, you should have valid license for GrapeCity Documents for PDF.
            //Create a pdf document.
            GrapeCity.Documents.Pdf.GcPdfDocument doc = new GrapeCity.Documents.Pdf.GcPdfDocument();
            // This page will save datas for multiple pages.
            GrapeCity.Documents.Pdf.Page page = doc.NewPage();

            //Create a PrintManager.
            GrapeCity.Documents.Excel.PrintManager printManager = new GrapeCity.Documents.Excel.PrintManager();

            //Get the pagination information of the workbook.
            IList<PageInfo> pages = printManager.Paginate(workbook);

            //Divide the multiple pages into 1 rows and 2 columns and printed them on one page.
            printManager.Draw(page, pages, 1, 2);

            //Save the document into pdf file.
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
                return "Multiple sheets one page.xlsx";
            }
        }
        public override string[] UsedResources
        {
            get
            {
                return new string[] { "xlsx\\Multiple sheets one page.xlsx" };
            }
        }
    }
}
