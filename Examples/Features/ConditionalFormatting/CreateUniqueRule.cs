using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.ConditionalFormatting
{
    public class CreateUniqueRule : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            object[,] data = new object[,]{
                {"Name", "City", "Birthday", "Eye color", "Weight", "Height"},
                {"Richard", "New York", new DateTime(1968, 6, 8), "Blue", 80, 165},
                {"Nia", "New York", new DateTime(1972, 7, 3), "Brown", 72, 134},
                {"Jared", "New York", new DateTime(1964, 3, 2), "Hazel", 71, 180},
                {"Natalie", "Washington", new DateTime(1972, 8, 8), "Blue", 80, 163},
                {"Damon", "Washington", new DateTime(1986, 2, 2), "Hazel", 58, 176},
                {"Angela", "Washington", new DateTime(1993, 2, 15), "Brown", 71, 145}
            };
            worksheet.Range["A1:F7"].Value = data;

            //Unique rule.
            IUniqueValues condition = worksheet.Range["E2:E7"].FormatConditions.AddUniqueValues();
            condition.DupeUnique = DupeUnique.Unique;
            condition.Font.Name = "Arial";
            condition.Interior.Color = Color.Pink;
        }
    }
}
