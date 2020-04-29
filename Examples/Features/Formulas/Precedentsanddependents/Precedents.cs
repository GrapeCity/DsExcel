using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Formulas.Precedentsanddependents
{
    public class Precedents : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Range["E2"].Formula = "=sum(A1:A2, B4,C1:C3)";
            worksheet.Range["A1"].Value = 1;
            worksheet.Range["A2"].Value = 2;
            worksheet.Range["B4"].Value = 3;
            worksheet.Range["C1"].Value = 4;
            worksheet.Range["C2"].Value = 5;
            worksheet.Range["C3"].Value = 6;

            foreach (var item in worksheet.Range["E2"].GetPrecedents())
            {
                item.Interior.Color = Color.Pink;
            }
        }
    }
}
