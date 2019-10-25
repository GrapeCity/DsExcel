using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting.PrintManager
{
    public class CustomPageInfos : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook, MemoryStream outputStream)
        {
            Stream fileStream = this.GetResourceStream("xlsx\\KeepTogether.xlsx");
            workbook.Open(fileStream);

            IWorksheet worksheet = workbook.Worksheets[0];

            //Firstly, create a printManager.
            GrapeCity.Documents.Excel.PrintManager printManager = new GrapeCity.Documents.Excel.PrintManager();

            //Get the natural pagination information of the worksheet.
            //The first page of the natural pagination is "A1:F37", the second page is from row "A38:F73" 
            IList<PageInfo> pages = printManager.Paginate(worksheet);

            //Custom the pageInfos.
            pages[0].PageContent.Range = worksheet.Range["A1:F36"]; // The first page is "A1:F36".
            pages[0].PageSettings.CenterHeader = "&KFF0000&20 Budget summary report"; // The center header of the first page will show the text "Budget summary report".
            pages[0].PageSettings.CenterFooter = "&KFF0000&20 Page &P"; // The center footer of the first page will show the page number "1".
            pages[1].PageContent.Range = worksheet.Range["A37:F73"]; // The second page is "A37:F73".

            //Save the modified pages into pdf file.
            printManager.SavePDF(outputStream, pages);
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
                return "KeepTogether.xlsx";
            }
        }
        public override string[] UsedResources
        {
            get
            {
                return new string[] { "xlsx\\KeepTogether.xlsx" };
            }
        }
    } 
}
