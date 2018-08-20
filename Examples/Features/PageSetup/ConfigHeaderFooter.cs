using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PageSetup
{
    public class ConfigHeaderFooter : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            var fileStream = this.GetTemplateStream();
            workbook.Open(fileStream);
            IWorksheet worksheet = workbook.Worksheets[0];

            //Set page headerfooter
            worksheet.PageSetup.LeftHeader = "&\"Arial,Italic\"LeftHeader";
            worksheet.PageSetup.CenterHeader = "&P";

            //Set page headerfooter's graphic
            worksheet.PageSetup.CenterFooter = "&G";
            System.IO.Stream stream = this.GetResourceStream("logo.png");
            worksheet.PageSetup.CenterFooterPicture.SetGraphicStream(stream, Drawing.ImageType.PNG);

            //If you have picture resources locally, you can also set graphic in this way.
            //worksheet.PageSetup.CenterFooter = "&G";
            //worksheet.PageSetup.CenterFooterPicture.Filename = @"C:\picture.png";

        }
        public override string TemplateName
        {
            get
            {
                return "PageSetup Demo.xlsx";
            }
        }
    }
}
