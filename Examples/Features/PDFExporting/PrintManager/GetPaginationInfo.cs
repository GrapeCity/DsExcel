using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting.PrintManager
{
    public class GetPaginationInfo : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook, MemoryStream outputStream)
        {
            var fileStream = this.GetResourceStream("xlsx\\Medical office start-up expenses 1.xlsx");
            workbook.Open(fileStream);

            IWorksheet worksheet = workbook.Worksheets[1];

            GrapeCity.Documents.Excel.PrintManager printManager = new GrapeCity.Documents.Excel.PrintManager();

            // The columnIndexs is [4, 12, 20], this means that the horizontal direction is split after the column 5th, 13th, and 21th. 
            IList<int> columnIndexs = printManager.GetPaginationInfo(worksheet, PaginationOrientation.Horizontal);
            // The rowIndexs is [42, 61], this means that the vertical direction is split after the row 43th and 62th.
            IList<int> rowIndexs = printManager.GetPaginationInfo(worksheet, PaginationOrientation.Vertical);

            // Save the pages into pdf file.
            IList<PageInfo> pages = printManager.Paginate(worksheet);
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

        public override string[] UsedResources
        {
            get
            {
                return new string[] { "xlsx\\Medical office start-up expenses 1.xlsx" };
            }
        }

        public override bool IsNew
        {
            get
            {
                return true;
            }
        }
    }
}
