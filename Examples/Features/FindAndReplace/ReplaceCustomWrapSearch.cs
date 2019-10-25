namespace GrapeCity.Documents.Excel.Examples.Features.FindAndReplace
{
    public class ReplaceCustomWrapSearch : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            // Prepare data
            worksheet.Range["A1:A8"].Value = new[] {
                "Whats new in GcExcel v2 sp2", "Render Excel ranges inside PDF in .NET Core",
                "Control pagination when printing Excel document to PDF in .NET Core (Support Team)",
                "How to format Pivot table styles in .NET Core (Support Team)",
                "Controlling page breaks when editing Excel files in .NET Core (Support Team)",
                "Combine different workbooks into PDF in .NET Core (Support Team)",
                "Repeating Excel rows/columns on exporting to PDF in .NET Core (Support Team)", "Using GcExcel with Kotlin"
            };

            // Find ".NET Core" and replace them with ".NET 5", starting after A4
            var what = ".NET Core";
            var replacement = ".NET 5";
            FindOptions settings = new FindOptions();
            var target = worksheet.UsedRange;
            var after = worksheet.Range["A4"];

            // Search start after A4
            IRange cellToReplace = after;
            do
            {
                cellToReplace = target.Find(what, cellToReplace, settings);
                if (cellToReplace == null)
                {
                    break;
                }

                // Replace
                cellToReplace.Value = cellToReplace.Text.Replace(what, replacement);
            } while (true);

            // Search reached the bottom of the range.
            // Wrap search start at the top-left corner.
            if (after != null)
            {
                do
                {
                    cellToReplace = target.Find(what, cellToReplace, settings);
                    if (cellToReplace == null)
                    {
                        break;
                    }

                    // Replace
                    cellToReplace.Value = cellToReplace.Text.Replace(what, replacement);

                    if (cellToReplace.Row == after.Row && cellToReplace.Column == after.Column)
                    {
                        break;
                    }
                } while (true);
            }
        }
    }

}
