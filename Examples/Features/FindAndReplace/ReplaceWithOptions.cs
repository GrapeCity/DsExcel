namespace GrapeCity.Documents.Excel.Examples.Features.FindAndReplace
{
    public class ReplaceWithOptions : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            // Prepare data

            // Skew matrix generator		
            // Input:
            // DegX    135     
            // DegY    45      
            // 
            // Output:
            // M11 1	    M12	1
            // M21 -1	M22	1
            // M31 0	    M32	0
            worksheet.Range["B1"].Value = "Skew matrix generator";
            worksheet.Range["A2:A4"].Value = new[] { "Input:", "DegX", "DegY" };
            worksheet.Range["B3"].Value = 135;
            worksheet.Range["B4"].Value = 45;
            worksheet.Range["A6"].Value = "Output:";
            worksheet.Range["A7:A9"].Value = new[] { "M11", "M21", "M31" };
            worksheet.Range["B7"].Value = 1;
            worksheet.Range["B8"].Formula = "=TAN(B3/180*3.14)";
            worksheet.Range["B9"].Value = 0;
            worksheet.Range["C7:C9"].Value = new[] { "M12", "M22", "M32" };
            worksheet.Range["D7"].Formula = "=TAN(B4/180*3.14)";
            worksheet.Range["D8"].Value = 1;
            worksheet.Range["D9"].Value = 0;

            // Replace 3.14 with PI()
            var searchRange = worksheet.UsedRange;
            searchRange.Replace(3.14, "PI()");

            // Replace M with m (Match case)
            searchRange.Replace("M", "m", new ReplaceOptions { MatchCase = true });

            // Replace m11 with M11 (Match whole word, match byte)
            searchRange.Replace("m11", "M11", new ReplaceOptions
            {
                LookAt = LookAt.Whole,
                MatchByte = true
            });

        }
    }

}
