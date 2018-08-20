using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Sparklines
{
    public class UseDateAxis : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            object[,] data = new object[,]
            {
                {"Number", "Date", "Customer", "Description", "Trend", "0-30 Days", "30-60 Days", "60-90 Days", ">90 Days", "Amount"},
                {"1001", new DateTime(2017, 5, 21), "Customer A", "Invoice 1001", null, 1200.15, 1916.18, 1105.23, 1806.53, null},
                {"1002", new DateTime(2017, 3, 18), "Customer B", "Invoice 1002", null, 896.23, 1005.53, 1800.56, 1150.49, null},
                {"1003", new DateTime(2017, 6, 15), "Customer C", "Invoice 1003", null, 827.63, 1009.23, 1869.23, 1002.56, null}
            };

            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Range["B2:K5"].Value = data;
            worksheet.Range["B:K"].ColumnWidth = 15;

            worksheet.Tables.Add(worksheet.Range["B2:K5"], true);
            worksheet.Tables[0].Columns[9].DataBodyRange.Formula = "=SUM(Table1[@[0-30 Days]:[>90 Days]])";

            //create a new group of sparklines.
            worksheet.Range["F3:F5"].SparklineGroups.Add(SparkType.Line, "G3:J5");

            worksheet.Range["G7:J7"].Value = new object[] { new DateTime(2011, 12, 16), new DateTime(2011, 12, 17), new DateTime(2011, 12, 18), new DateTime(2011, 12, 19) };
            worksheet.Range["F3"].SparklineGroups[0].DateRange = "G7:J7";
            worksheet.Range["F3"].SparklineGroups[0].Axes.Horizontal.Axis.Visible = true;
            worksheet.Range["F3"].SparklineGroups[0].Axes.Horizontal.Axis.Color.Color = Color.Green;
        }
    }
}
