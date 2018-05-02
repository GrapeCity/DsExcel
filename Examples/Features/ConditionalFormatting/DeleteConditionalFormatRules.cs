using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.ConditionalFormatting
{
    public class DeleteConditionalFormatRules : ExampleBase
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

            //iconset rule
            IIconSetCondition iconset = worksheet.Range["E2:E7"].FormatConditions.AddIconSetCondition();
            iconset.IconSet = workbook.IconSets[IconSetType.Icon3TrafficLights1];

            //cell value rule added later, it has the highest priority, set StopIfTrue to true, if cell match condition, it will not apply icon set rule.
            IFormatCondition cellvalueRule = worksheet.Range["E2:E7"].FormatConditions.Add(FormatConditionType.CellValue, FormatConditionOperator.Between, "71", "79") as IFormatCondition;
            cellvalueRule.StopIfTrue = true;

            //delete icon set rule.
            (worksheet.Range["E2:E7"].FormatConditions[1] as IIconSetCondition).Delete();

            //delete all the rules
            worksheet.Range["E2:E7"].FormatConditions.Delete();
        }
    }
}
