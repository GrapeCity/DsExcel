using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PageSetup
{
    public class ConfigPrintTailColumns : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            var fileStream = this.GetResourceStream("xlsx\\RepeatTails.xlsx");
            workbook.Open(fileStream);
            IWorksheet worksheet = workbook.Worksheets[0];

            //Repeat the column "I" at the right of each page when saving pdf.
            worksheet.PageSetup.PrintTailColumns = "$I:$I";
        }

        public override string TemplateName
        {
            get
            {
                return "RepeatTails.xlsx";
            }
        }
        public override string[] UsedResources
        {
            get
            {
                return new string[] { "xlsx\\RepeatTails.xlsx" };
            }
        }

        public override bool ShowViewer
        {
            get
            {
                return false;
            }
        }
    }
}
