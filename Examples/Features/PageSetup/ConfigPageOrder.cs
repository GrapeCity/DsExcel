using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PageSetup
{
    public class ConfigPageOrder : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            var fileStream = this.GetResourceStream("xlsx\\PageSetup Demo.xlsx");
            workbook.Open(fileStream);
            IWorksheet worksheet = workbook.Worksheets[0];

            //Set page order, default is DownThenOver.
            worksheet.PageSetup.Order = Order.OverThenDown;
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
