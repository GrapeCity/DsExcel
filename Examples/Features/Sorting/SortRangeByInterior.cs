using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Sorting
{
    public class SortRangeByInterior : ExampleBase
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

            worksheet.Range["F2"].Interior.Color = Color.LightPink;
            worksheet.Range["F3"].Interior.Color = Color.LightGreen;
            worksheet.Range["F4"].Interior.Color = Color.LightPink;
            worksheet.Range["F5"].Interior.Color = Color.LightGreen;
            worksheet.Range["F6"].Interior.Color = Color.LightBlue;
            worksheet.Range["F7"].Interior.Color = Color.LightPink;
            
            //"F4" will in the top.
            worksheet.Sort.SortFields.Add(new CellColorSortField(worksheet.Range["F2:F7"], worksheet.Range["F4"].DisplayFormat.Interior, SortOrder.Ascending));
            worksheet.Sort.Range = worksheet.Range["A2:F7"];
            worksheet.Sort.Orientation = SortOrientation.Columns;
            worksheet.Sort.Apply();
        }
    }
}
