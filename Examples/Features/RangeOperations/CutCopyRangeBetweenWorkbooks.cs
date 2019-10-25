using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.RangeOperations
{
    public class CutCopyRangeBetweenWorkbooks : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //Load template file Home inventory.xlsx from resource
            GrapeCity.Documents.Excel.Workbook source_workbook = new GrapeCity.Documents.Excel.Workbook();
            var source_fileStream = this.GetResourceStream("xlsx\\Home inventory.xlsx");
            source_workbook.Open(source_fileStream);

            //Hide gridline
            workbook.ActiveSheet.SheetView.DisplayGridlines = false;

            workbook.ActiveSheet.Range["A1"].Value = "Copy content from the first sheet of source workbook";
            workbook.ActiveSheet.Range["A1"].Font.Color = Color.Red;
            workbook.ActiveSheet.Range["A1"].Font.Bold = true;

            //Copy content of active sheet from source workbook to the current sheet at A2
            source_workbook.ActiveSheet.GetUsedRange().Copy(workbook.ActiveSheet.Range["A2"], PasteType.Default | PasteType.RowHeights | PasteType.ColumnWidths);

            workbook.ActiveSheet.Range["C21"].Value = "Cut content from the second sheet of source workbook";
            workbook.ActiveSheet.Range["C21"].Font.Color = Color.Red;
            workbook.ActiveSheet.Range["C21"].Font.Bold = true;

            //Cut content of second sheet from source workbook to the current sheet at C22
            source_workbook.Worksheets[1].Range["2:15"].Cut(workbook.ActiveSheet.Range["C22"]);

            //Make the theme of two workbooks same
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
