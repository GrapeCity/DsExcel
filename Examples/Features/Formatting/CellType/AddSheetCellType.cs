using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Formatting.CellType
{
    public class AddSheetCellType : ExampleBase
    {
        public override void Execute(Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            ButtonCellType buttonCellType = new ButtonCellType
            {
                Text = "Button",
                ButtonBackColor = "Azure",
                MarginLeft = 10,
                MarginRight = 10
            };

            worksheet.CellType = buttonCellType;
        }
    }
}
