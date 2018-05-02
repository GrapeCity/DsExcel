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
                {"Name", "City", "Birthday", "Sex", "Weight", "Height"},
                {"Bob", "NewYork", new DateTime(1968, 6, 8), "male", 80, 180},
                {"Betty", "NewYork", new DateTime(1972, 7, 3), "female", 72, 168},
                {"Gary", "NewYork", new DateTime(1964, 3, 2), "male", 71, 179},
                {"Hunk", "Washington", new DateTime(1972, 8, 8), "male", 80, 171},
                {"Cherry", "Washington", new DateTime(1986, 2, 2), "female", 58, 161},
                { "Eva", "Washington", new DateTime(1993, 2, 15), "female", 71, 180}
            };
            worksheet.Range["A1:F7"].Value = data;

            //icon set rule.
            IIconSetCondition condition = worksheet.Range["E2:E7"].FormatConditions.AddIconSetCondition();
            condition.IconSet = workbook.IconSets[IconSetType.Icon3Symbols];
            condition.IconCriteria[1].Operator = FormatConditionOperator.GreaterEqual;
            condition.IconCriteria[1].Value = 50;
            condition.IconCriteria[1].Type = ConditionValueTypes.Percent;
            condition.IconCriteria[2].Operator = FormatConditionOperator.GreaterEqual;
            condition.IconCriteria[2].Value = 70;
            condition.IconCriteria[2].Type = ConditionValueTypes.Percent;
        }
    }
}
