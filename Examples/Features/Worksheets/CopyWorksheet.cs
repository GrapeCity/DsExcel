using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Worksheets
{
    public class CopyWorksheet : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //Load template file AgingReport.xlsx from resource
            var fileStream = this.GetResourceStream("xlsx\\AgingReport.xlsx");
            workbook.Open(fileStream);

            //Copy the active sheet to the end of current workbook
            var copy_worksheet = workbook.ActiveSheet.Copy();
            copy_worksheet.Name = "Copy of " + workbook.ActiveSheet.Name;
        }

        public override string TemplateName
        {
            get
            {
                return "AgingReport.xlsx";
            }
        }

        public override string[] UsedResources
        {
            get
            {
                return new string[] { "xlsx\\AgingReport.xlsx" };
            }
        }
    }
}
