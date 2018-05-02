using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.RangeOperations
{
    public class InsertCellsDeleteCells : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet1 = workbook.Worksheets[0];
            IWorksheet worksheet2 = workbook.Worksheets.Add();

            var data = new object[,]
            {
                {1, 2, 3},
                {4, 5, 6},
                {7, 8, 9}
            };

            worksheet1.Range["A1:C3"].Value = data;
            worksheet2.Range["A1:C3"].Value = data;

            //Insert cells
            worksheet1.Range["A2"].Insert();//Auto
            worksheet1.Range["B2"].Insert(InsertShiftDirection.Down);
            worksheet1.Range["C2"].Insert(InsertShiftDirection.Right);

            //Delete cells
            worksheet2.Range["A2"].Delete();//Auto
            worksheet2.Range["B2"].Delete(DeleteShiftDirection.Left);
            worksheet2.Range["C2"].Delete(DeleteShiftDirection.Up);
        }

    }
}
