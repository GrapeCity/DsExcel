using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PageSetup
{
    public class ConfigSheetPrintSettings : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            var fileStream = this.GetTemplateStream("PageSetup Demo.xlsx");
            workbook.Open(fileStream);
            IWorksheet worksheet = workbook.Worksheets[0];

            //Set sheet
            worksheet.PageSetup.PrintGridlines = true;
            worksheet.PageSetup.PrintHeadings = true;
            worksheet.PageSetup.BlackAndWhite = true;
            worksheet.PageSetup.PrintComments = PrintLocation.InPlace;
            worksheet.PageSetup.PrintErrors = PrintErrors.Dash;

        }
    }
}
