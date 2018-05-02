using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.ConditionalFormatting
{
    public class CreateCellValueRule : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Range["B:C"].ColumnWidth = 17;

            object[,] data = new object[,]{
                {"Name", "City", "Birthday", "Eye color", "Weight", "Height"},
                {"Richard", "New York", new DateTime(1968, 6, 8), "Blue", 67, 165},
                {"Nia", "New York", new DateTime(1972, 7, 3), "Brown", 62, 134},
                {"Jared", "New York", new DateTime(1964, 3, 2), "Hazel", 72, 180},
                {"Natalie", "Washington", new DateTime(1972, 8, 8), "Blue", 66, 163},
                {"Damon", "Washington", new DateTime(1986, 2, 2), "Hazel", 76, 176},
                {"Angela", "Washington", new DateTime(1993, 2, 15), "Brown", 68, 145}
            };
            worksheet.Range["A1:F7"].Value = data;

            //weight between 71 and 79, set its interior color to yellow.
            IFormatCondition condition = worksheet.Range["E2:E7"].FormatConditions.Add(FormatConditionType.CellValue, FormatConditionOperator.Between, 66, 70) as IFormatCondition;
            condition.Interior.Color = Color.LightGreen;
        }
    }
}
