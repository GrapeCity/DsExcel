namespace GrapeCity.Documents.Excel.Examples.Features.FindAndReplace
{
    public class FindMatchByte : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            // This option is valid when culture is ja-JP or zh-CN.
            workbook.Culture = System.Globalization.CultureInfo.GetCultureInfo("ja-JP");
            IWorksheet worksheet = workbook.Worksheets[0];

            // Prepare data
            worksheet.Range["A1:A4"].Value = new[] {
                "Mario Games", "スーパーマリオブラザーズ",
                "ﾏﾘｵ&ﾙｲｰｼﾞRPG3 DX", "マリオ＆ルイージRPG1 DX"
            };

            // Find the first cell that contains "マリオ" (match width) 
            // and mark it with red foreground.
            IRange searchRange = worksheet.UsedRange;
            FindOptions matchByteOptions = new FindOptions { MatchByte = true };
            var marioCell = searchRange.Find("マリオ", null, matchByteOptions);
            marioCell.Font.Color = System.Drawing.Color.Red;

            // Find the first cell that contains "ルイージ" (ignore width) 
            // and mark it with green background.
            var luigiCell = searchRange.Find("ルイージ");
            luigiCell.Interior.Color = System.Drawing.Color.Green;

        }
    }

}
