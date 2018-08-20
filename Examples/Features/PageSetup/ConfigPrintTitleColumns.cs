using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PageSetup
{
    public class ConfigPrintTitleColumns : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            var fileStream = this.GetTemplateStream();
            workbook.Open(fileStream);
            IWorksheet worksheet = workbook.Worksheets[0];

            //Set columns to repeat at left
            worksheet.PageSetup.PrintTitleColumns = "$D:$G";

        }

        public override string TemplateName
        {
            get
            {
                return "PageSetup Print Title Columns.xlsx";
            }
        }
    }
}
