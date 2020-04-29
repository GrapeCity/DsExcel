using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Formatting.CellType
{

    public class AddRowColumnCellType : ExampleBase
    {
        public override void Execute(Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Columns[3].ColumnWidthInPixel = 100;
            CheckBoxCellType cellType = new CheckBoxCellType
            {
                Caption = "CheckBox",
                TextTrue = "True",
                TextFalse = "False",
                IsThreeState = true,
                TextAlign = CheckBoxAlign.Right
            };
            worksheet.Columns[3].CellType = cellType;
            worksheet.Range["D1:D10"].Value = true;

            ButtonCellType buttonCellType = new ButtonCellType
            {
                Text = "Button",
                ButtonBackColor = "Azure",
                MarginLeft = 10,
                MarginRight = 10
            };

            worksheet.Rows[3].CellType = buttonCellType;
        }
    }
}
