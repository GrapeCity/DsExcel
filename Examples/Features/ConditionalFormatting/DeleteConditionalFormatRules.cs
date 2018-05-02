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
                {"Name", "City", "Birthday", "Eye color", "Weight", "Height"},
                {"Richard", "New York", new DateTime(1968, 6, 8), "Blue", 67, 165},
                {"Nia", "New York", new DateTime(1972, 7, 3), "Brown", 62, 134},
                {"Jared", "New York", new DateTime(1964, 3, 2), "Hazel", 72, 180},
                {"Natalie", "Washington", new DateTime(1972, 8, 8), "Blue", 66, 163},
                {"Damon", "Washington", new DateTime(1986, 2, 2), "Hazel", 76, 176},
                {"Angela", "Washington", new DateTime(1993, 2, 15), "Brown", 68, 145}
            };
            worksheet.Range["A1:F7"].Value = data;

            //iconset rule
            IIconSetCondition iconset = worksheet.Range["E2:E7"].FormatConditions.AddIconSetCondition();
            iconset.IconSet = workbook.IconSets[IconSetType.Icon3TrafficLights1];

            //cell value rule added later, it has the highest priority, set StopIfTrue to true, if cell match condition, it will not apply icon set rule.
            IFormatCondition cellvalueRule = worksheet.Range["E2:E7"].FormatConditions.Add(FormatConditionType.CellValue, FormatConditionOperator.Between, "66", "70") as IFormatCondition;
            cellvalueRule.StopIfTrue = true;

            //delete icon set rule.
            (worksheet.Range["E2:E7"].FormatConditions[1] as IIconSetCondition).Delete();

            //delete all the rules
            worksheet.Range["E2:E7"].FormatConditions.Delete();
        }
    }
}
