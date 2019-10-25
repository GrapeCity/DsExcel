using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Worksheets
{
    public class MoveWorksheetBetweenWorkbooks : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //Load template file Home inventory.xlsx from resource to the source workbook
            GrapeCity.Documents.Excel.Workbook source_workbook = new GrapeCity.Documents.Excel.Workbook();
            var source_fileStream = this.GetResourceStream("xlsx\\Home inventory.xlsx");
            source_workbook.Open(source_fileStream);

            //Move content of active sheet from source workbook to the current workbook before the first sheet
            var move_worksheet = source_workbook.ActiveSheet.MoveBefore(workbook.Worksheets[0]);
            move_worksheet.Name = "Move of Home inventory";
            move_worksheet.Activate();

            workbook.Theme = source_workbook.Theme;
        }

        public override string TemplateName
        {
            get
            {
                return "Home inventory.xlsx";
            }
        }

        public override string[] UsedResources
        {
            get
            {
                return new string[] { "xlsx\\Home inventory.xlsx" };
            }
        }
    }
}
