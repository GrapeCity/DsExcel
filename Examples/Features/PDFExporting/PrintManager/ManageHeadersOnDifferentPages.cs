using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting.PrintManager
{
    public class ManageHeadersOnDifferentPages : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook, MemoryStream outputStream)
        {
            Stream fileStream = this.GetResourceStream("xlsx\\MultipleHeaders.xlsx");
            workbook.Open(fileStream);
            IWorksheet worksheet = workbook.Worksheets[0];

            IList<RepeatSetting> repeatSettings = new List<RepeatSetting>();

            //The title rows of the "B2:F87" is "$2:$2"
            RepeatSetting repeatSetting = new RepeatSetting
            {
                TitleRowStart = 1,
                TitleRowEnd = 1,
                Range = worksheet.Range["B2:F87"]
            };
            repeatSettings.Add(repeatSetting);

            //The title rows of the "B91:F189" is "$91:$91"
            RepeatSetting repeatSetting2 = new RepeatSetting
            {
                TitleRowStart = 88,
                TitleRowEnd = 88,
                Range = worksheet.Range["B89:F149"]
            };
            repeatSettings.Add(repeatSetting2);

            //Create a PrintManager.
            GrapeCity.Documents.Excel.PrintManager printManager = new GrapeCity.Documents.Excel.PrintManager();

            //Get the pagination information of the worksheet.
            IList<PageInfo> pages = printManager.Paginate(worksheet, null, repeatSettings);

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
                return "MultipleHeaders.xlsx";
            }
        }
        public override string[] UsedResources
        {
            get
            {
                return new string[] { "xlsx\\MultipleHeaders.xlsx" };
            }
        }
    }
}
