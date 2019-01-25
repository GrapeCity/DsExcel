using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PageSetup
{
    public class ConfigFirstPageNumber : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            var fileStream = this.GetResourceStream("xlsx\\PageSetup Demo.xlsx");
            workbook.Open(fileStream);
            IWorksheet worksheet = workbook.Worksheets[0];

            worksheet.PageSetup.CenterFooter = "&P";

            //Set first page number, default is p1.
            worksheet.PageSetup.FirstPageNumber = 3;

        }

        public override string TemplateName
        {
            get
            {
                return "PageSetup Demo.xlsx";
            }
        }

        public override string[] UsedResources
        {
            get
            {
                return new string[] { "xlsx\\PageSetup Demo.xlsx", "logo.png" };
            }
        }
    }
}
