namespace GrapeCity.Documents.Excel.Examples.Features.Grouping
{
    public class AddSubtotalAdvanced : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IRange targetRange = workbook.ActiveSheet.Range["A1:C9"];
            // Set data
            targetRange.Value = new object[,]
            {
                {"Grade", "Class", "Score", "Student ID"},
                {1, 1, 93, 1},
                {1, 1, 87, 2},
                {1, 2, 97, 3},
                {1, 2, 95, 4},
                {2, 1, 83, 5},
                {2, 1, 87, 6},
                {2, 2, 96, 7},
                {2, 2, 83, 8}
            };

            // Group by Grade select Average(Score)
            targetRange.Subtotal(groupBy: 1, // Grade
                subtotalFunction: ConsolidationFunction.Average,
                totalList: new[] { 3 }, // Score
                replace: false,
                pageBreaks: true);

            // Group by Class select Average(Score)
            targetRange.Subtotal(groupBy: 2, // Class
                subtotalFunction: ConsolidationFunction.Average,
                totalList: new[] { 3 }, // Score
                replace: false);

            workbook.ActiveSheet.Range["C:C"].NumberFormat = "0;0;0;@";

            targetRange.AutoFit();
        }
    }
}
