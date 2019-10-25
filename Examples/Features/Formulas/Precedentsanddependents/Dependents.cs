using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Formulas.Precedentsanddependents
{
    public class Dependents : ExampleBase
    {
        public override bool IsNew => true;
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Range["A1"].Value = 100;
            worksheet.Range["C1"].Formula = "=$A$1";
            worksheet.Range["E1:E5"].Formula = "=$A$1";
            foreach (var item in worksheet.Range["A1"].GetDependents())
            {
                item.Interior.Color = Color.Azure;
            }
        }
    }
}
