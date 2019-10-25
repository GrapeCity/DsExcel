namespace GrapeCity.Documents.Excel.Examples.Features.FindAndReplace
{
    public class FindLookAtSearchOrderDirectionMatchCase : ExampleBase
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
            worksheet.Range["B3"].Formula = "=DATE(2019,5,1)";
            worksheet.Range["B3"].NumberFormat = "yyyy-mm-dd;@";
            worksheet.Range["C3"].Formula = "=B3+1";
            worksheet.Range["C3"].NumberFormat = "yyyy-mm-dd;@";
            worksheet.UsedRange.AutoFit();

            var searchRange = worksheet.Range["A1:C3"];

            // Find the last occurrence of 1 in text (match whole word, backward and by columns)
            // and mark it with blue foreground and bigger font 
            var lastValue1 = searchRange.Find(1, null, new FindOptions
            {
                LookIn = FindLookIn.Values,
                SearchDirection = SearchDirection.Previous,
                LookAt = LookAt.Whole,
                SearchOrder = SearchOrder.ByColumns
            });
            lastValue1.Font.Color = System.Drawing.Color.Blue;
            lastValue1.Font.Size += 8;

        }
    }

}
