namespace GrapeCity.Documents.Excel.Examples.Features.FindAndReplace
{
    public class FindWithLookIn : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            // Prepare data

            // Add day to date		
            // Day    Date	    Result
            // 1      2019-05-01	2019-05-02
            worksheet.Range["A2:C2"].Value = new[] { "Day", "Date", "Result" };
            worksheet.Range["A1"].Value = "Add day to date";
            worksheet.Range["A3"].Value = 1;
            worksheet.Range["A3"].AddComment("Enter the day offset");
            worksheet.Range["B3"].Formula = "=DATE(2019,5,1)";
            worksheet.Range["B3"].NumberFormat = "yyyy-mm-dd;@";
            worksheet.Range["C3"].Formula = "=B3+1";
            worksheet.Range["C3"].NumberFormat = "yyyy-mm-dd;@";
            worksheet.UsedRange.AutoFit();

            // Find the first occurrence of "2019" in the formula bar 
            // and mark it with green foreground color
            var searchRange = worksheet.Range["A1:C3"];
            var first2019InFormulaBar = searchRange.Find("2019", null,
                new FindOptions { LookIn = FindLookIn.Formulas });
            first2019InFormulaBar.Font.Color = System.Drawing.Color.Green;

            // Find the first occurrence of 1 in text
            // and mark it with blue foreground 
            var firstValue1 = searchRange.Find(1, null, 
                new FindOptions { LookIn = FindLookIn.Values });
            firstValue1.Font.Color = System.Drawing.Color.Blue;

            // Find the first occurrence of "day" in comments
            // and mark it with yellow background 
            var firstDayComments = searchRange.Find("day", null,
                new FindOptions { LookIn = FindLookIn.Comments });
            firstDayComments.Interior.Color = System.Drawing.Color.Yellow;

            // Find the last occurrence of "2019" in the formula property
            // and mark it with purple foreground
            var last2019OnlyFormula = searchRange.Find("2019",
                options: new FindOptions
            {
                LookIn = FindLookIn.OnlyFormulas,
                SearchDirection = SearchDirection.Previous
            });
            last2019OnlyFormula.Font.Color = System.Drawing.Color.Purple;
        }
    }

}
