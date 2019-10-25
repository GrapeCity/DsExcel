using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting.PrintManager
{
    public class PrintSpecificPages : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook, MemoryStream outputStream)
        {
            Stream fileStream = this.GetResourceStream("xlsx\\PrintSpecificPDFPages.xlsx");
            workbook.Open(fileStream);

            //Firstly, create a printManager.
            GrapeCity.Documents.Excel.PrintManager printManager = new GrapeCity.Documents.Excel.PrintManager();

            //Get the natural pagination information of the workbook.
            IList<PageInfo> pages = printManager.Paginate(workbook);

            //Pick some pages to print.
            IList<PageInfo> newPages = new List<PageInfo>();
            newPages.Add(pages[0]);
            newPages.Add(pages[2]);

            //Update the page number and the page settings of each page. The page number is continuous.
            printManager.UpdatePageNumberAndPageSettings(newPages);

            //Save the pages into pdf file.
            printManager.SavePDF(outputStream, newPages);
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
                return "PrintSpecificPDFPages.xlsx";
            }
        }
        public override string[] UsedResources
        {
            get
            {
                return new string[] { "xlsx\\PrintSpecificPDFPages.xlsx" };
            }
        }
    }
}
