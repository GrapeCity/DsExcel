using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.RangeOperations
{
    public class InsertDeleteRowsColumns : ExampleBase
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

            //Insert rows
            worksheet1.Range["A2:B2"].EntireRow.Insert();
            worksheet1.Range["3:5"].Insert();

            //Insert columns
            worksheet1.Range["B3:B5"].EntireColumn.Insert();
            worksheet1.Range["A:A"].Insert();

            //Delete rows
            worksheet2.Range["A3:A5"].EntireRow.Delete();
            worksheet2.Range["2:4"].Delete();

            //Delete columns
            worksheet2.Range["B3:B5"].EntireColumn.Delete();
            worksheet2.Range["A:A"].Delete();
        }
    }
}
