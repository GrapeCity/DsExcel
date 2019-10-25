using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Formatting.CellType
{
    public class AddButtonCellType : ExampleBase
    {
        public override bool IsNew => true;
        public override void Execute(Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            ButtonCellType cellType = new ButtonCellType
            {
                Text = "Hello",
                ButtonBackColor = "Azure",
                MarginLeft = 10,
                MarginRight = 10
            };

            worksheet.Range["C5"].CellType = cellType;
        }
    }
}
