using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Formatting.CellType
{
    public class AddComboBoxCellType : ExampleBase
    {
        public override bool IsNew => true;
        public override void Execute(Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            ComboBoxCellType cellType = new ComboBoxCellType
            {
                EditorValueType = EditorValueType.Value
            };
            ComboBoxCellItem item = new ComboBoxCellItem
            {
                Value = "US",
                Text = "United States"
            };
            cellType.Items.Add(item);
            item = new ComboBoxCellItem
            {
                Value = "CN",
                Text = "China"
            };
            cellType.Items.Add(item);
            item = new ComboBoxCellItem
            {
                Value = "JP",
                Text = "Japan"
            };
            cellType.Items.Add(item);

            worksheet.Range["C5"].CellType = cellType;
            worksheet.Range["C5"].Value = "CN";
        }
    }
}
