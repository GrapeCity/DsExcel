using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Sorting
{
    public class SortRangeByFontColor : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            object[,] data = new object[,]{
                {"Name", "City", "Birthday", "Eye color", "Weight", "Height"},
                {"Richard", "New York", new DateTime(1968, 6, 8), "Blue", 67, 165},
                {"Nia", "New York", new DateTime(1972, 7, 3), "Brown", 62, 134},
                {"Jared", "New York", new DateTime(1964, 3, 2), "Hazel", 72, 180},
                {"Natalie", "Washington", new DateTime(1972, 8, 8), "Blue", 66, 163},
                {"Damon", "Washington", new DateTime(1986, 2, 2), "Hazel", 76, 176},
                {"Angela", "Washington", new DateTime(1993, 2, 15), "Brown", 68, 145}
            };

            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Range["A1:F7"].Value = data;
            worksheet.Range["A:F"].ColumnWidth = 15;

            worksheet.Range["F2"].Font.Color = Color.Red;
            worksheet.Range["F3"].Font.Color = Color.Gainsboro;
            worksheet.Range["F4"].Font.Color = Color.Green;
            worksheet.Range["F5"].Font.Color = Color.Gainsboro;
            worksheet.Range["F6"].Font.Color = Color.Gainsboro;
            worksheet.Range["F7"].Font.Color = Color.Gainsboro;

            //The font colors which are same to "F7" cell's color will in the top.
            worksheet.Sort.SortFields.Add(new FontColorSortField(worksheet.Range["F2:F7"], worksheet.Range["F7"].Font.Color, SortOrder.Ascending));
            worksheet.Sort.Range = worksheet.Range["A2:F7"];
            worksheet.Sort.Orientation = SortOrientation.Columns;
            worksheet.Sort.Apply();
        }
    }
}
