using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.FindAndReplace
{
    public class FindBasicUsage : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            // Prepare data
            const string CorrectWord = "Macro";
            worksheet.Range["A1:D5"].Value = CorrectWord;

            const string MisspelledWord = "marco";
            worksheet.Range["A2,C3,D1"].Value = MisspelledWord;

            // Find the first misspelled word
            IRange searchRange = worksheet.Range["A1:D5"];
            IRange firstMisspelled = searchRange.Find(MisspelledWord);

            // Mark it with red foreground
            firstMisspelled.Font.Color = System.Drawing.Color.Red;
        }
    }

}
