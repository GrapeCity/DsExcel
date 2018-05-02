using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Tables
{
    public class PerformCalculationsInTable : ExampleBase
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

            //add table.
            worksheet.Tables.Add(worksheet.Range["A1:F7"], true);

            //"Name" column's total row calculation function will be "=SUBTOTAL(103,[Name])"
            worksheet.Tables[0].Columns[0].TotalsCalculation = TotalsCalculation.Count;
            //"Weight" column's total row calculation function will be "=SUBTOTAL(101,[Weight])"
            worksheet.Tables[0].Columns[4].TotalsCalculation = TotalsCalculation.Average;
            worksheet.Tables[0].ShowTotals = true;
        }
    }
}
