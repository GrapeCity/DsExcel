using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PageSetup
{
    public class ConfigPrintTitleRows : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            var fileStream = this.GetTemplateStream();
            workbook.Open(fileStream);
            IWorksheet worksheet = workbook.Worksheets[0];

            //Set rows to repeat at top
            worksheet.PageSetup.PrintTitleRows = "$5:$10";

        }

        public override string TemplateName
        {
            get
            {
                return "PageSetup Print Title Rows.xlsx";
            }
        }
    }
}
