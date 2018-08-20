using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.RangeOperations
{
    public class CellInfo : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            // cell's value B2
            string cell = GrapeCity.Documents.Excel.CellInfo.CellIndexToName(1, 1);
            worksheet.Range[cell].Interior.Color = Color.LightBlue;

            int rowIndex, columnIndex;
            // rowIndex is 3 and columnIndex is 2
            GrapeCity.Documents.Excel.CellInfo.CellNameToIndex("C4", out rowIndex, out columnIndex);
            worksheet.Range[rowIndex, columnIndex].Interior.Color = Color.LightCoral;

            // column is D
            string column = GrapeCity.Documents.Excel.CellInfo.ColumnIndexToName(3);
            worksheet.Range[String.Format("{0}:{0}", column)].Interior.Color = Color.LightGreen;

            // columnIndex is 4
            columnIndex = GrapeCity.Documents.Excel.CellInfo.ColumnNameToIndex("E");
            worksheet.Columns[columnIndex].Interior.Color = Color.LightSalmon;

            // row is 3
            string row = GrapeCity.Documents.Excel.CellInfo.RowIndexToName(2);
            worksheet.Range[String.Format("{0}:{0}", row)].Interior.Color = Color.LightSteelBlue;

            // rowIndex is 4
            rowIndex = GrapeCity.Documents.Excel.CellInfo.RowNameToIndex("5");
            worksheet.Rows[rowIndex].Interior.Color = Color.LightSkyBlue;
        }

        public override bool IsNew
        {
            get
            {
                return true;
            }
        }
    }
}
