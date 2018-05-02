using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Formulas
{
    public class DirtyAndCalculation : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Range["A1"].Value = 1;
            worksheet.Range["A2"].Formula = "=A1";
            worksheet.Range["A3"].Formula = "=SUM(A1, A2)";

            //when get value, calc engine will first calculate and cache the result, then returns the cached result.
            var value_A2 = worksheet.Range["A2"].Value;
            var value_A3 = worksheet.Range["A3"].Value;

            //disable calc engine.
            workbook.EnableCalculation = false;

            //Dirty() method will clear the cached value of the workbook.
            workbook.Dirty();
            //Calculate() will not work, because of workbook.EnablCalculation is false.
            workbook.Calculate();
            //it returns 0 because of no cache value exist.
            var value_A2_1 = worksheet.Range["A2"].Value;
            var value_A3_1 = worksheet.Range["A3"].Value;

            worksheet.Range["A1"].Value = 2;
            //enable calc engine.
            workbook.EnableCalculation = true;
            //Dirty() method will clear the cached value of Range A2:A3.
            worksheet.Range["A2:A3"].Dirty();
            //Calculate() method will calculate and cache the result, it will return the cache value directly when get value later.
            worksheet.Range["A2:A3"].Calculate();

            //it returns cache value directly, does not calculate again.
            var value_A2_2 = worksheet.Range["A2"].Value;
            var value_A3_2 = worksheet.Range["A3"].Value;

        }
    }
}
