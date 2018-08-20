using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.ConditionalFormatting
{
    public class CreateTwoColorScaleRule : ExampleBase
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

            //two color scale.
            IColorScale twoColorScaleRule = worksheet.Range["E2:E7"].FormatConditions.AddColorScale(ColorScaleType.TwoColorScale);
            twoColorScaleRule.ColorScaleCriteria[0].Type = ConditionValueTypes.Number;
            twoColorScaleRule.ColorScaleCriteria[0].Value = 62;
            twoColorScaleRule.ColorScaleCriteria[0].FormatColor.Color = Color.Red;

            twoColorScaleRule.ColorScaleCriteria[1].Type = ConditionValueTypes.Number;
            twoColorScaleRule.ColorScaleCriteria[1].Value = 76;
            twoColorScaleRule.ColorScaleCriteria[1].FormatColor.Color = Color.Green;
        }

    }
}
