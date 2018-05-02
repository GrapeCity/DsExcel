using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Formulas
{
    public class UseArrayFormula : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Range["E4:J5"].Value = new object[,]
             {
                {1, 2, 3},
                {4, 5, 6}
             };

            worksheet.Range["I6:J8"].Value = new object[,]
            {
                {2, 2},
                {3, 3},
                {4, 4}
            };

            //O     P      Q
            //2     4      #N/A
            //12    15     #N/A
            //#N/A  #N/A   #N/A
            worksheet.Range["O9:Q11"].FormulaArray = "=E4:G5*I6:J8";

            //judge if Range O9 has array formula.
            if (worksheet.Range["O9"].HasArray)
            {
                //set O9's entire array's interior color.
                var currentarray = worksheet.Range["O9"].CurrentArray;
                currentarray.Interior.Color = Color.Green;
            }
        }
    }
}
