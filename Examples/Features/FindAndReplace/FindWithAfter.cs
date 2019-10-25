namespace GrapeCity.Documents.Excel.Examples.Features.FindAndReplace
{
    public class FindWithAfter : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            // Prepare data
            const string CorrectWord = "Macro";
            worksheet.Range["A1:D5"].Value = CorrectWord;

            const string MisspelledWord = "marco";
            worksheet.Range["A2,C3,D1"].Value = MisspelledWord;

            // Find all misspelled words and mark them with red background
            IRange searchRange = worksheet.Range["A1:D5"];
            IRange misspelledCell = null;
            do
            {
                misspelledCell = searchRange.Find(MisspelledWord, misspelledCell);
                if (misspelledCell == null)
                {
                    break;
                }
                misspelledCell.Interior.Color = System.Drawing.Color.Red;
            } while (true);
        }
    }

}
