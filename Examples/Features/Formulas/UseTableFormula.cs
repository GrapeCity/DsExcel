using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Formulas
{
    public class UseTableFormula : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Range["A:E"].ColumnWidth = 15;
            worksheet.Range["A1:E3"].Value = new object[,]
            {
                {"SalesPerson", "Region",  "SalesAmount", "ComPct", "ComAmt"},
                {"Joe", "North", 260, 0.10, null},
                {"Nia", "South", 660, 0.15, null},
            };

            worksheet.Tables.Add(worksheet.Range["A1:E3"], true);
            worksheet.Tables[0].Name = "DeptSales";
            worksheet.Tables[0].Columns["ComPct"].DataBodyRange.NumberFormat = "0%";

            //Use table formula in table range.
            worksheet.Tables[0].Columns["ComAmt"].DataBodyRange.Formula = "=[@ComPct]*[@SalesAmount]";

            //Use table formula out of table range.
            worksheet.Range["F2"].Formula = "=SUM(DeptSales[@SalesAmount])";
            worksheet.Range["G2"].Formula = "=SUM(DeptSales[[#Data],[SalesAmount]])";
            worksheet.Range["H2"].Formula = "=SUM(DeptSales[SalesAmount])";
            worksheet.Range["I2"].Formula = "=SUM(DeptSales[@ComPct], DeptSales[@ComAmt])";

            //judge if Range F2:I2 have formula.
            for (int i = 5; i <= 8; i++)
            {
                if (worksheet.Range[1, i].HasFormula)
                {
                    worksheet.Range[1, i].Interior.Color = Color.LightBlue;
                }
            }
        }
    }
}
