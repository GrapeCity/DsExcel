using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PageSetup
{
    public class ConfigPaperScaling : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            var fileStream = this.GetTemplateStream("PageSetup Demo.xlsx");
            workbook.Open(fileStream);
            IWorksheet worksheet = workbook.Worksheets[0];

            //Set paper scaling
            //Method 1: Set percent scale 
            worksheet.PageSetup.IsPercentScale = true;
            worksheet.PageSetup.Zoom = 150;

            //Or Method 2: Fit to page's wide & tall
            //worksheet.PageSetup.IsPercentScale = false;
            //worksheet.PageSetup.FitToPagesWide = 3;
            //worksheet.PageSetup.FitToPagesTall = 4;

        }
    }
}
