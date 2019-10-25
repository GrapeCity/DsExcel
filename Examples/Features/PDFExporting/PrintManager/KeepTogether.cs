using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting.PrintManager
{
    public class KeepTogether : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook, MemoryStream outputStream)
        {
            Stream fileStream = this.GetResourceStream("xlsx\\KeepTogether.xlsx");
            workbook.Open(fileStream);
            IWorksheet worksheet = workbook.Worksheets[0];

            //The first page of the natural pagination is from row 1th to 37th, the second page is from row 38th to 73th. 
            IList<IRange> keepTogetherRanges = new List<IRange>();
            //The row 37th and 38th need to keep together. So the pagination results are: the first page if from row 1th to 36th, the second page is from row 37th to 73th.
            keepTogetherRanges.Add(worksheet.Range["37:38"]);

            //Create a PrintManager.
            GrapeCity.Documents.Excel.PrintManager printManager = new GrapeCity.Documents.Excel.PrintManager();

            //Get the pagination information of the worksheet.
            IList<PageInfo> pages = printManager.Paginate(worksheet, keepTogetherRanges, null);

            //Save the pages into pdf file.
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
