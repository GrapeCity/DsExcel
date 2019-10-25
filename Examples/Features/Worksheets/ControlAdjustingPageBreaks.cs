using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Worksheets
{
    public class ControlAdjustingPageBreaks : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet sheet = workbook.Worksheets[0];

            sheet.Range["A1:E5"].Value = new object[,]
            {
                {1, 2, 3, 4, 5},
                {6, 7, 8, 9, 10},
                {11, 12, 13, 14, 15},
                {16, 17, 18, 19, 20},
                { 21, 22, 23, 24, 25},
            };

            //Add page break
            sheet.HPageBreaks.Add(sheet.Range["D4"]); //add a horizontal page break before the fourth row.
            sheet.VPageBreaks.Add(sheet.Range["D4"]); //add a vertical page break before the fourth column.

            //delete rows and columns before the page breaks, the page breaks will be adjusted.
            sheet.Range["1:1"].Delete(); // the hPageBreak is before the third row.
            sheet.Range["A:A"].Delete(); // the vPageBreak is before the third column.

            //set the page breaks are fixed, it will not be adjusted when inserting/deleting rows/columns.
            sheet.FixedPageBreaks = true;

            sheet.Range["1:1"].Delete(); // the hPageBreak is still before the third row.
            sheet.Range["A:A"].Delete(); // the vPageBreak is still before the third column.
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
