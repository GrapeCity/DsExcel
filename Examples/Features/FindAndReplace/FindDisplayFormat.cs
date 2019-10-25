namespace GrapeCity.Documents.Excel.Examples.Features.FindAndReplace
{
    public class FindDisplayFormat : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            // Prepare data
            worksheet.Range["A1:C3"].Value="Text";

            var b2 = worksheet.Range["B2"];
            b2.Interior.Color = System.Drawing.Color.Red;
            b2.Font.Color = System.Drawing.Color.White;
            b2.Value = "B2";

            var a2 = worksheet.Range["A2"];
            a2.Interior.Color = System.Drawing.Color.Orange;
            a2.Font.Color = System.Drawing.Color.White;
            a2.Value = "A2";

            // Find cells with red background and white foreground,
            // and highlight them with bold and bigger text

            // Create a temporary sheet to build a IDisplayFormat
            IWorksheet displayFormatFactoryWorksheet = workbook.Worksheets.Add();
            IRange displayFormatFactoryRange = displayFormatFactoryWorksheet.Range["A1"];
            displayFormatFactoryRange.Interior.Color = System.Drawing.Color.Red;
            displayFormatFactoryRange.Font.Color = System.Drawing.Color.White;
            IDisplayFormat searchFormat = displayFormatFactoryRange.DisplayFormat;

            // Find the first occurrence
            IRange searchRange = worksheet.UsedRange;
            FindOptions options = new FindOptions { SearchFormat = searchFormat };
            IRange foundCell = searchRange.Find("*", null, options);

            // Highlight the found range
            foundCell.Font.Bold = true;
            foundCell.Font.Size += 8;

            // Dispose the temporary sheet
            displayFormatFactoryWorksheet.Delete();
        }
    }

}
