using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Filtering.DateFiltering
{
    public class DateFilter : ExampleBase
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
            worksheet.Range["A:F"].ColumnWidth = 15;

            var criteria1 = new DateTime(1972, 7, 3).ToString();
            var criteria2 = new DateTime(1993, 2, 15).ToString();
            //filter date between 1972.7.3 and 1993.2.15
            worksheet.Range["A1:F7"].AutoFilter(2, ">=" + criteria1, AutoFilterOperator.And, "<=" + criteria2);
        }
    }
}
