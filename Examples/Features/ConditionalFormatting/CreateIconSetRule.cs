using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.ConditionalFormatting
{
    public class CreateIconSetRule : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

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

            //icon set rule.
            IIconSetCondition condition = worksheet.Range["E2:E7"].FormatConditions.AddIconSetCondition();
            condition.IconSet = workbook.IconSets[IconSetType.Icon3Symbols];
            condition.IconCriteria[1].Operator = FormatConditionOperator.GreaterEqual;
            condition.IconCriteria[1].Value = 30;
            condition.IconCriteria[1].Type = ConditionValueTypes.Percent;
            condition.IconCriteria[2].Operator = FormatConditionOperator.GreaterEqual;
            condition.IconCriteria[2].Value = 70;
            condition.IconCriteria[2].Type = ConditionValueTypes.Percent;
        }
    }
}
