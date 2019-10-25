namespace GrapeCity.Documents.Excel.Examples.Features.Grouping
{
    public class AddSubtotal : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IRange targetRange = workbook.ActiveSheet.Range["A1:C9"];
            // Set data
            targetRange.Value = new object[,]
            {
                {"Player", "Side", "Commander"},
                {1, "Soviet", "AI"},
                {2, "Soviet", "AI"},
                {3, "Soviet", "Human"},
                {4, "Allied", "Human"},
                {5, "Allied", "Human"},
                {6, "Allied", "AI"},
                {7, "Empire", "AI"},
                {8, "Empire", "AI"}
            };

            // Subtotal
            targetRange.Subtotal(groupBy: 2, // Side
                subtotalFunction: ConsolidationFunction.Count,
                totalList: new[] { 2 } // Side
                );

            targetRange.AutoFit();
        }
        public override bool IsNew
        {
            get
            {
                return true;
            }
        }
    }

}
