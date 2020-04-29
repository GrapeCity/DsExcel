using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Formatting.CellType
{
    public class AddCheckBoxCellType : ExampleBase
    {
        public override void Execute(Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            CheckBoxCellType cellType = new CheckBoxCellType
            {
                Caption = "Caption",
                TextTrue = "True",
                TextFalse = "False",
                TextIndeterminate = "Indeterminate",
                IsThreeState = true,
                TextAlign = CheckBoxAlign.Right
            };

            worksheet.Range["C5:C6"].CellType = cellType;
            worksheet.Range["C5"].Value = true;
            worksheet.Range["C6"].Value = false;
        }
    }
}
