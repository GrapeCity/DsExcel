using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PageSetup
{
    public class ConfigPageCenter : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            var fileStream = this.GetResourceStream("xlsx\\PageSetup Demo.xlsx");
            workbook.Open(fileStream);
            IWorksheet worksheet = workbook.Worksheets[0];

            //Set center, default is false.
            worksheet.PageSetup.CenterHorizontally = true;
            worksheet.PageSetup.CenterVertically = true;
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
                return new string[] { "xlsx\\PageSetup Demo.xlsx" };
            }
        }
    }
}
