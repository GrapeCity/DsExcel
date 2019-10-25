using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PageSetup
{
    public class ConfigPrintTailRows : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            var fileStream = this.GetResourceStream("xlsx\\RepeatTails.xlsx");
            workbook.Open(fileStream);
            IWorksheet worksheet = workbook.Worksheets[1];

            //Repeat the row 50th at the bottom of each page when saving pdf.
            worksheet.PageSetup.PrintTailRows = "$50:$50";
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
