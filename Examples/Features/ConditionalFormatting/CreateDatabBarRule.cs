using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.ConditionalFormatting
{
    public class CreateDatabBarRule : ExampleBase
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

            //data bar rule.
            IDataBar dataBar = worksheet.Range["E2:E7"].FormatConditions.AddDatabar();

            dataBar.MinPoint.Type = ConditionValueTypes.LowestValue;
            dataBar.MinPoint.Value = null;
            dataBar.MaxPoint.Type = ConditionValueTypes.HighestValue;
            dataBar.MaxPoint.Value = null;

            dataBar.BarFillType = DataBarFillType.Gradient;
            dataBar.BarColor.Color = Color.Green;
            dataBar.Direction = DataBarDirection.Context;
            dataBar.AxisColor.Color = Color.Red;
            dataBar.AxisPosition = DataBarAxisPosition.Automatic;
            dataBar.NegativeBarFormat.BorderColorType = DataBarNegativeColorType.Color;
            dataBar.NegativeBarFormat.BorderColor.Color = Color.Blue;
            dataBar.NegativeBarFormat.ColorType = DataBarNegativeColorType.Color;
            dataBar.NegativeBarFormat.Color.Color = Color.Pink;
            dataBar.ShowValue = false;
        }
    }
}
