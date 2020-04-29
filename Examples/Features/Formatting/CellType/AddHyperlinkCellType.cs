using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Formatting.CellType
{
    public class AddHyperlinkCellType : ExampleBase
    {
        public override void Execute(Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            HyperLinkCellType cellType = new HyperLinkCellType
            {
                Text = "Google",
                LinkColor = "Blue",
                LinkToolTip = "Search by google",
                VisitedLinkColor = "Green",
                Target = HyperLinkTargetType.Blank
            };

            worksheet.Range["C5"].CellType = cellType;
            worksheet.Range["C5"].Value = "http://www.google.com";
        }
    }
}
