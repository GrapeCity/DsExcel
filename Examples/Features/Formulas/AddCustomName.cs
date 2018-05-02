using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Formulas
{
    public class AddCustomName : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet1 = workbook.Worksheets[0];
            IWorksheet worksheet2 = workbook.Worksheets.Add();

            worksheet1.Range["C8"].NumberFormat = "0.0000";

            worksheet1.Names.Add("test1", "=Sheet1!$A$1");
            worksheet1.Names.Add("test2", "=Sheet1!test1*2");
            workbook.Names.Add("test3", "=Sheet1!$A$1");

            worksheet1.Range["A1"].Value = 1;

            //C6's value is 1.
            worksheet1.Range["C6"].Formula = "=test1";
            //C7's value is 3.
            worksheet1.Range["C7"].Formula = "=test1 + test2";
            //C8's value is 6.283185307
            worksheet1.Range["C8"].Formula = "=test2*PI()";

            //judge if Range C6:C8 have formula.
            for (int i = 5; i <= 7; i++)
            {
                if (worksheet1.Range[i, 2].HasFormula)
                {
                    worksheet1.Range[i, 2].Interior.Color = Color.LightBlue;
                }
            }

            //worksheet1 range A2's value is 1.
            worksheet2.Range["A2"].Formula = "=test3";
            //judge if Range A2 has formula.
            if (worksheet2.Range["A2"].HasFormula)
            {
                worksheet2.Range["A2"].Interior.Color = Color.LightBlue;
            }

            //set r1c1 formula.
            worksheet2.Range["A3"].FormulaR1C1 = "=R[-1]C";
        }
    }
}
