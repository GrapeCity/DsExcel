namespace GrapeCity.Documents.Excel.Examples.Features.FindAndReplace
{
    public class ReplaceBasicUsage : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            // Prepare data
            worksheet.Range["A1:A3"].Value = new[] {
                "Render Excel ranges inside PDF in .NET Core",
                "Control pagination when printing Excel document to PDF in .NET Core (Support Team)",
                "How to format Pivot table styles in .NET Core (Support Team)"
            };

            // Replace ".NET Core" with ".NET 5"
            worksheet.UsedRange.Replace(".NET Core", ".NET 5");
        }
    }

}
